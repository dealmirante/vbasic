VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Begin VB.Form frmADMSectorSup 
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin FPSpread.vaSpread sprPv 
      Height          =   4005
      Left            =   45
      OleObjectBlob   =   "frmADMSectorSup.frx":0000
      TabIndex        =   3
      Top             =   735
      Width           =   4950
   End
   Begin VB.CommandButton botSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3615
      TabIndex        =   1
      Top             =   4995
      Width           =   1335
   End
   Begin VB.CommandButton botSalvar 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   45
      TabIndex        =   2
      Top             =   4995
      Width           =   1335
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
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   4860
   End
End
Attribute VB_Name = "frmADMSectorSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub botSalir_Click()
Unload Me
End Sub

Private Sub botSalvar_Click()
Dim loc As String
Dim m2 As Variant, sec As Variant
Dim sql As String
Dim res As Boolean

Aplicacion.ComienzoTrans
    
loc = frmAdmSectorProveedor.TreeView1.SelectedItem.Key
res = True

For i = 1 To sprPv.MaxRows
    sprPv.GetText 3, i, m2
    sprPv.GetText 1, i, sec
    sql = " update estadis.z_sector_local set metros_cuadrados = " & m2
    sql = sql & " Where cod_local = '" & loc & "' "
    sql = sql & " And cod_sector = " & sec
    
    res = Aplicacion.EjecutarDAO(sql)
    If Not res Then
        Exit For
    End If
Next

If res Then
    Aplicacion.TerminarConExitoTrans
    MsgBox "Modificación exitosa ", vbOKOnly + vbExclamation, "Atención"
Else
    Aplicacion.TerminarConErrorTrans
    MsgBox "Error en la modificación ", vbOKOnly + vbCritical, "Error"
End If

End Sub

Private Sub Form_Load()

Me.Top = frmAdmSectorProveedor.Top + 1300
Me.Left = 950

lblSecI.caption = frmAdmSectorProveedor.lblLocalSec

L_CargaGrilla

End Sub

Private Sub L_CargaGrilla()
Dim rs As Recordset

sprPv.MaxRows = 0

If Aplicacion.ObtenerRsDAO(frmAdmSectorProveedor.sqlSector, rs) Then
    Do While Not rs.EOF
        sprPv.MaxRows = sprPv.MaxRows + 1
        
        sprPv.SetText 1, sprPv.MaxRows, Trim(rs!cod_sector)
        sprPv.SetText 2, sprPv.MaxRows, Trim(rs!Descrip)
        sprPv.SetText 4, sprPv.MaxRows, str(rs!m2)
    rs.MoveNext
    Loop
    rs.Close
End If


End Sub
