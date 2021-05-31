VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmADMSectorProv 
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   5010
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5580
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton botModif 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   315
      TabIndex        =   19
      Top             =   5070
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton botSalvar 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   315
      TabIndex        =   14
      Top             =   5055
      Width           =   1335
   End
   Begin VB.CommandButton botSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3405
      TabIndex        =   13
      Top             =   5055
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   285
      TabIndex        =   1
      Top             =   720
      Width           =   4485
      Begin MSMask.MaskEdBox mskPartVenta 
         Height          =   345
         Left            =   1980
         TabIndex        =   15
         Top             =   2535
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.ListBox lstSRub 
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   1785
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ListBox lstRub 
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ListBox lstProv 
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox cboSubRubro 
         Height          =   315
         Left            =   1965
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1785
         Width           =   2175
      End
      Begin VB.ComboBox cboRubro 
         Height          =   315
         Left            =   1965
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cboProveedor 
         Height          =   315
         Left            =   1965
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Peso de venta (%)"
         Height          =   375
         Index           =   3
         Left            =   300
         TabIndex        =   5
         Top             =   2535
         Width           =   1575
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sub Rubro"
         Height          =   375
         Index           =   2
         Left            =   285
         TabIndex        =   4
         Top             =   1785
         Width           =   1575
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rubro"
         Height          =   375
         Index           =   1
         Left            =   285
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor"
         Height          =   375
         Index           =   0
         Left            =   285
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame frSec 
      Height          =   990
      Left            =   270
      TabIndex        =   16
      Top             =   4005
      Visible         =   0   'False
      Width           =   4515
      Begin VB.ComboBox cboSector 
         Height          =   315
         Left            =   1935
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   315
         Width           =   2175
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sectores"
         Height          =   375
         Index           =   4
         Left            =   285
         TabIndex        =   17
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Nota: El porcentaje de ocupación de proveedor en el sector permanecerá en cero hasta que el usuario lo modifique"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   345
      TabIndex        =   9
      Top             =   4290
      Width           =   4365
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
Attribute VB_Name = "frmADMSectorProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim vLoc As String
Dim vSector As String
Dim vSectorOrig As String

Dim vSecuencia As Integer
Dim vProv As String
Dim vDescrip As String

Dim nodo As Node

Dim tipo As String
Public Sub altas(n As Node, ByRef pProv As String, ByRef pDescrip As String, ByRef pSec As Integer)

vLoc = Mid(n.Key, 1, 3)
vSector = Mid(n.Key, InStr(1, n.Key, "*") + 1)

vSecuencia = 1

Set nodo = n
tipo = "ALTA"

Me.Show 1

pProv = vProv
pDescrip = vDescrip
pSec = vSecuencia

End Sub

Public Sub modificacion(n As Node, NuevoSec As Integer, pSector As Integer, pSec As Integer)

vLoc = Mid(n.Key, 1, 3)
If NuevoSec = -1 Then
    vSector = Mid(n.Key, InStr(1, n.Key, "*") + 1, InStr(1, n.Key, "#") - InStr(1, n.Key, "*") - 1)
    vSectorOrig = vSector
Else
    vSectorOrig = Mid(n.Key, InStr(1, n.Key, "*") + 1, InStr(1, n.Key, "#") - InStr(1, n.Key, "*") - 1)
    vSector = NuevoSec
End If
vProv = Mid(n.Key, InStr(1, n.Key, "#") + 1, InStr(1, n.Key, "@") - InStr(1, n.Key, "#") - 1)
vSecuencia = Mid(n.Key, InStr(1, n.Key, "@") + 1)

Set nodo = n
tipo = "MOD"

Me.Show 1

pProv = vProv
pDescrip = vDescrip
pSec = vSecuencia
pSector = vSector

End Sub



Private Function L_TesteoInexistenciaProv() As Boolean
Dim sql As String
Dim sqlSec As String
Dim rs As Recordset
Dim existe As Boolean
Dim cadena As String
Dim strsec As String

sql = " Select cod_sector,cod_prov , cod_rubr, cod_srub "
sql = sql & " From estadis.Z_SECTOR_LOCAL_PROV "
sql = sql & " Where cod_local = '" & vLoc & "' "
sql = sql & " And cod_prov = '" & lstProv.List(cboProveedor.ListIndex) & "' "

sqlSec = "  Select nvl(max(secuencia),0) descrip "
sqlSec = sqlSec & " From estadis.Z_SECTOR_LOCAL_PROV "
sqlSec = sqlSec & " Where cod_local = '" & vLoc & "' "
sqlSec = sqlSec & " and cod_sector =  " & vSector

Func.Func_ObtenerDesc sqlSec, strsec


cadena = ""
existe = False
If Aplicacion.ObtenerRsDAO(sql, rs) Then
    Do While Not rs.EOF
        If rs!cod_prov = lstProv.List(cboProveedor.ListIndex) Then
           If rs!cod_rubr = cboRubro.Text Then
           'Si el proveedor - rubro exise
                If IsNull(rs!cod_srub) Or cboSubRubro.ListIndex = -1 Then
                  If rs!cod_sector = vSector Then
                  'No se permite tener repetido el dato en el mismo sector
                    cadena = cadena & "Se ha detectado :" & cboProveedor.Text & "-" & cboRubro.Text _
                    & " para el mismo sector "
                    existe = True
                    Exit Do
                  Else
                   cadena = cadena & "Se ha detectado :" & cboProveedor.Text & "-" & cboRubro.Text _
                   & " para el sector : " & vLoc & "-" & rs!cod_sector & Chr(10) & Chr(13)
                  End If
                ElseIf rs!cod_srub = lstSRub.List(cboSubRubro.ListIndex) Then
                  If rs!cod_sector = vSector Then
                  'No se permite tener repetido el dato en el mismo sector
                    cadena = cadena & "Se ha detectado :" & cboProveedor.Text & "-" & cboRubro.Text & " - " & cboSubRubro.Text _
                    & " para el mismo sector "
                    existe = True
                    Exit Do
                  Else
                   cadena = cadena & "Se ha detectado :" & cboProveedor.Text & "-" & cboRubro.Text & " - " & cboSubRubro.Text _
                   & " para el sector : " & vLoc & "-" & rs!cod_sector & Chr(10) & Chr(13)
                   
                  End If
                Else
                    vSecuencia = Val(strsec) + 1
                End If

           End If
        End If
    rs.MoveNext
    Loop
End If
If existe Then
    MsgBox cadena, vbYes, "Error"
    L_TesteoInexistenciaProv = False
Else
If cadena <> "" Then
   If MsgBox(cadena & " Confirma ? ", vbYesNo, "Atencion") = vbYes Then
      L_TesteoInexistenciaProv = True
      MsgBox "Recuerde modificar los porcentajes de ventas", vbOKOnly, "Atención"
   Else
      L_TesteoInexistenciaProv = False
   End If
Else
L_TesteoInexistenciaProv = True
End If
End If
End Function



Private Sub botModif_Click()
Dim sql As String


'Primero se borra el prov
sql = "Delete from estadis.z_sector_local_prov Where cod_sector = " & vSectorOrig
sql = sql & " And cod_local = '" & vLoc & "' "
sql = sql & " And cod_prov  = '" & vProv & "' "
sql = sql & " And secuencia = " & vSecuencia

Aplicacion.ComienzoTrans

If Aplicacion.EjecutarDAO(sql) Then
    If L_TesteoInexistenciaProv() Then
        sql = " Insert into estadis.Z_SECTOR_LOCAL_PROV ("
        sql = sql & " SECUENCIA,"
        sql = sql & " COD_SECTOR,"
        sql = sql & " COD_LOCAL,"
        sql = sql & " COD_PROV,"
        sql = sql & " COD_RUBR,"
        sql = sql & " COD_SRUB,"
        sql = sql & " PARTICIPACION_VENTA,"
        sql = sql & " PARTICIPACION_OCUPA )"
        sql = sql & " Values ( "
        sql = sql & vSecuencia & ", "
        sql = sql & vSector & ", "
        sql = sql & "'" & vLoc & "', "
        sql = sql & "'" & lstProv.List(cboProveedor.ListIndex) & "', "
        sql = sql & "'" & cboRubro.Text & "', "
        sql = sql & "'" & lstSRub.List(cboSubRubro.ListIndex) & "', "
        sql = sql & mskPartVenta.Text & ", 0 ) "

        If Aplicacion.EjecutarDAO(sql) Then
           Aplicacion.TerminarConExitoTrans
           MsgBox "El dato fue cargado con exito ", vbExclamation, "Atención"
           
        Else
           Aplicacion.TerminarConErrorTrans
           MsgBox "No se pudo asignar el sector ", vbCritical, "Atención"
           vSector = -1
        End If
    Else
        Aplicacion.TerminarConErrorTrans
        vSector = -1
    End If
Else
    MsgBox "No se puede modificar el registro", vbCritical + vbOKOnly, "Error"


End If

    Unload Me
End Sub

Private Sub botSalir_Click()
vProv = "*"
vSector = -1
Unload Me
End Sub

Private Sub botSalvar_Click()
Dim sql As String

If L_TesteoInexistenciaProv() Then
    
    sql = " Insert into estadis.Z_SECTOR_LOCAL_PROV ("
    sql = sql & " SECUENCIA,"
    sql = sql & " COD_SECTOR,"
    sql = sql & " COD_LOCAL,"
    sql = sql & " COD_PROV,"
    sql = sql & " COD_RUBR,"
    sql = sql & " COD_SRUB,"
    sql = sql & " PARTICIPACION_VENTA,"
    sql = sql & " PARTICIPACION_OCUPA )"
    sql = sql & " Values ( "
    sql = sql & vSecuencia & ", "
    sql = sql & vSector & ", "
    sql = sql & "'" & vLoc & "', "
    sql = sql & "'" & lstProv.List(cboProveedor.ListIndex) & "', "
    sql = sql & "'" & cboRubro.Text & "', "
    sql = sql & "'" & lstSRub.List(cboSubRubro.ListIndex) & "', "
    sql = sql & mskPartVenta.Text & ", 0 ) "
    
    Aplicacion.ComienzoTrans
    
    If Aplicacion.EjecutarDAO(sql) Then
       Aplicacion.TerminarConExitoTrans
       MsgBox "El dato fue cargado con exito ", vbExclamation, "Atención"
       vProv = lstProv.List(cboProveedor.ListIndex)
       vDescrip = cboProveedor.Text
    Else
       Aplicacion.TerminarConErrorTrans
       MsgBox "No se pudo asignar el sector ", vbCritical, "Atención"
       vProv = "*"
    End If

    Unload Me

End If

End Sub

Private Sub cboProveedor_Click()
Dim sql As String

sql = " Select r.cod_rubr From baires.rubro r, baires.provrubr p"
sql = sql & " Where p.cod_rubr = r.cod_rubr And p.cod_prov = '" & lstProv.List(cboProveedor.ListIndex) & "' "

cboSubRubro.Clear
cboRubro.Clear

'Llenar el combo de rubro
FuncCbos_LlenarCbo cboRubro, sql


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

sql = " Select cod_srub,descr From baires.subrubro "
sql = sql & " Where cod_rubr = '" & cboRubro.Text & "' "

'Llenar el combo de subrubro
FuncCbos_LlenarCboLst cboSubRubro, lstSRub, sql

End Sub


Private Sub cboRubro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        cboRubro.ListIndex = -1
        cboSubRubro.Clear
    End If

End Sub


Private Sub cboSector_Click()
vSector = cboSector.ItemData(cboSector.ListIndex)
End Sub


Private Sub cboSubRubro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        cboSubRubro.ListIndex = -1
    End If

End Sub

Private Sub Form_Load()
Dim sql As String

Me.Top = frmAdmSectorProveedor.Top + 1300
Me.Left = 950

sql = "select cod_prov,descrip From  baires.proveedor order by descrip"

'Funcion que llena el combo de proveedor
FuncCbos_LlenarCboLst cboProveedor, lstProv, sql


sql = "select cod_sector,descrip From  estadis.z_sector Where cod_sector in (Select cod_sector From estadis.z_sector_local where cod_local = '" & vLoc & "')"

'Funcion que llena el combo de proveedor
FuncCbos.FuncCbos_LlenarCboiTEM cboSector, sql

If tipo = "ALTA" Then
    lblSecI.caption = nodo.Parent.Text & " - " & nodo.Text
Else
    lblSecI.caption = nodo.Parent.Parent.Text & " - " & nodo.Parent.Text
    L_SeteoDatos
End If

mskPartVenta.Text = 100

End Sub

Private Sub L_SeteoDatos()
Dim sql As String

    Func.Func_SetearCboItem cboSector, vSector

    Func.Func_SetearCboConLst cboProveedor, lstProv, vProv
    cboProveedor.Enabled = False
    
    Func.Func_SetearCboSTR cboRubro, frmAdmSectorProveedor.txtRubro
    Func.Func_SetearCboConLst cboSubRubro, lstSRub, frmAdmSectorProveedor.txtSubRubro

    mskPartVenta.Text = frmAdmSectorProveedor.txtVenta
    
    frSec.Visible = True
    botModif.Visible = True
    
End Sub

