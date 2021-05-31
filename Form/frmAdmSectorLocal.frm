VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmADMSectorLocal 
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   5010
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5580
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton botSalvar 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   285
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton botSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3405
      TabIndex        =   5
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   285
      TabIndex        =   1
      Top             =   720
      Width           =   4485
      Begin VB.ListBox lstProv 
         Height          =   255
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin MSMask.MaskEdBox mskSuperficie 
         Height          =   345
         Left            =   1980
         TabIndex        =   7
         Top             =   1665
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Format          =   "####.00"
         PromptChar      =   " "
      End
      Begin VB.ComboBox cboSector 
         Height          =   315
         Left            =   1965
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Superficie"
         Height          =   375
         Index           =   3
         Left            =   300
         TabIndex        =   3
         Top             =   1665
         Width           =   1575
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sector"
         Height          =   375
         Index           =   0
         Left            =   285
         TabIndex        =   2
         Top             =   960
         Width           =   1575
      End
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
      Height          =   495
      Left            =   405
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmADMSectorLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cod_local As String

Dim nD As String
Dim nS As Integer
Public Sub Altas(loc As String, ByRef pS As Integer, ByRef pD As String)
cod_local = loc

Me.Show 1

pS = nS
pD = nD

End Sub

Private Sub botSalir_Click()
nS = -1
Unload Me
End Sub



Private Sub botSalvar_Click()
Dim sql As String


sql = " Insert into estadis.z_sector_local (cod_sector,cod_local,metros_cuadrados) "
sql = sql & " Values ( "
sql = sql & cboSector.ItemData(cboSector.ListIndex) & ", "
sql = sql & "'" & cod_local & "' , "
sql = sql & mskSuperficie.Text & ") "


Aplicacion.ComienzoTrans

If Aplicacion.EjecutarDAO(sql) Then
    Aplicacion.TerminarConExitoTrans
    MsgBox "El dato fue cargado con exito ", vbExclamation, "Atención"
    nS = cboSector.ItemData(cboSector.ListIndex)
    nD = cboSector.Text
Else
    Aplicacion.TerminarConErrorTrans
    MsgBox "No se pudo asignar el sector ", vbCritical, "Atención"
    nS = -1
End If

Unload Me

End Sub

Private Sub cboSector_Change()
    nS = -1
    Unload Me
End Sub


Private Sub Form_Load()
Dim sql As String

Me.Top = frmAdmSectorProveedor.Top + 1300
Me.Left = 950

sql = "select cod_sector,descrip From  estadis.z_sector "

'Funcion que llena el combo de proveedor
FuncCbos.FuncCbos_LlenarCboiTEM cboSector, sql

lblSecI.caption = frmAdmSectorProveedor.lblLocal

mskSuperficie.Text = 0

End Sub
