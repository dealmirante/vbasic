VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Begin VB.Form frmADMSectorOC 
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
      OleObjectBlob   =   "frmADMSectorOC.frx":0000
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
      Left            =   60
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
      Height          =   570
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   4905
   End
End
Attribute VB_Name = "frmADMSectorOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub botSalir_Click()
Unload Me
End Sub

Private Sub botSalvar_Click()
Dim loc As String, vSector As Integer
Dim oc As Variant, sec As Variant, pv As Variant
Dim sql As String
Dim res As Boolean
Dim Total As Variant

sprPv.GetText 4, sprPv.MaxRows, Total

If Val(Total) = 100 Then
Aplicacion.ComienzoTrans
    
loc = frmAdmSectorProveedor.TreeView1.SelectedItem.Parent.Key
vSector = frmAdmSectorProveedor.txtCodSec

res = True

For i = 1 To sprPv.MaxRows - 1
    sprPv.GetText 4, i, oc
    sprPv.GetText 5, i, sec
    sprPv.GetText 1, i, pv
    
    sql = " update estadis.z_sector_local_prov set participacion_ocupa = " & oc
    sql = sql & " Where cod_local = '" & loc & "' "
    sql = sql & " And cod_sector = " & vSector & " "
    sql = sql & " And cod_prov = '" & pv & "' "
    sql = sql & " And secuencia = " & sec
    
    
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
Else
    MsgBox "Los porcentajes no suma 100% ", vbOKOnly + vbCritical, "Error"
End If
End Sub

Private Sub Form_Load()

Me.Top = frmAdmSectorProveedor.Top + 1300
Me.Left = 950

lblSecI.caption = frmAdmSectorProveedor.lblLocalSec & " - " & frmAdmSectorProveedor.lblSector

L_CargaGrilla

End Sub

Private Sub L_CargaGrilla()
Dim rs As Recordset

sprPv.MaxRows = 0

If Aplicacion.ObtenerRsDAO(frmAdmSectorProveedor.sqlProv, rs) Then
    Do While Not rs.EOF
        sprPv.MaxRows = sprPv.MaxRows + 1
        
        sprPv.SetText 1, sprPv.MaxRows, Trim(rs!cod_prov)
        sprPv.SetText 2, sprPv.MaxRows, Trim(rs!Descrip)
        sprPv.SetText 3, sprPv.MaxRows, Trim(rs!cod_rubr)
        sprPv.SetText 4, sprPv.MaxRows, str(rs!oc)
        sprPv.SetText 5, sprPv.MaxRows, str(rs!secuencia)
    rs.MoveNext
    Loop
    rs.Close
End If

Spread.Spread_TotalesGrillas sprPv, 3, 3

End Sub
