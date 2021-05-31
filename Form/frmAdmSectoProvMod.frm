VERSION 5.00
Begin VB.Form frmADMSectorProvMod 
   Caption         =   "Modificación de datos Proveedor"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton botSalvar 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton botSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4215
      Begin VB.ListBox lstSRub 
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   840
         Width           =   135
      End
      Begin VB.ListBox lstRub 
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   360
         Width           =   135
      End
      Begin VB.TextBox txtPartVenta 
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox cboSubRubro 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox cboRubro 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Peso de venta"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sub Rubro"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rubro"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Label lblProvI 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Nota: El porcentaje de ocupación permanecerá en cero hasta que el usuario lo modifique"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Label lblSecI 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmADMSectorProvMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub botSalir_Click()
Unload Me
End Sub

Private Sub cboProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        cboProveedor.ListIndex = -1
        cboRubro.Clear
        cboSubRubro.Clear
    End If

End Sub


Private Sub cboRubro_Click()
Dim sql As String

sql = " Select cod_srub,descrip From baires.subrubro "
sql = sql & " Where p.cod_rubr = '" & cboRubro.Text & "' "

'Llenar el combo de subrubro
'FuncCbos_LlenarCboLst cboSubRubro, lstSRub, sql
End Sub


Private Sub cboRubro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        cboRubro.ListIndex = -1
        cboSubRubro.Clear
    End If

End Sub


Private Sub Form_Load()
Dim sql As String

sql = " Select cod_rubr From baires.rubro r, baires.provrubr p"
sql = sql & " Where p.cod_rubr = r.cod_rubr And p.cod_prov = '" & cboProveedor.Text & "' "

cboSubRubro.Clear
cboRubro.Clear

'Llenar el combo de rubro
'FuncCbos_LlenarCboLst cboRubro, lstRub, sql


End Sub
