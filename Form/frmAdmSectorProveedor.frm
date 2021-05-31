VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmAdmSectorProveedor 
   Caption         =   "Aministración de sectores"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   10425
   Begin VB.CommandButton botArbol 
      Caption         =   "Armar Arbol"
      Height          =   405
      Left            =   1215
      TabIndex        =   34
      Top             =   5880
      Width           =   2745
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   5670
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   10001
      _Version        =   327680
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmAdmSectorProveedor.frx":0000
   End
   Begin VB.Frame frProv 
      Height          =   5685
      Left            =   5460
      TabIndex        =   12
      Top             =   45
      Visible         =   0   'False
      Width           =   4365
      Begin VB.TextBox txtDescripSR 
         Height          =   330
         Left            =   2325
         TabIndex        =   33
         Top             =   2610
         Width           =   1800
      End
      Begin VB.CommandButton botModProv 
         Caption         =   "Modificar Proveedor"
         Height          =   465
         Left            =   615
         TabIndex        =   32
         Top             =   5055
         Width           =   1290
      End
      Begin VB.CommandButton botDelProv 
         Caption         =   "Eliminar Proveedor"
         Height          =   465
         Left            =   2565
         TabIndex        =   31
         Top             =   5070
         Width           =   1290
      End
      Begin VB.TextBox txtPart 
         Height          =   345
         Left            =   1725
         TabIndex        =   24
         Top             =   3930
         Width           =   2430
      End
      Begin VB.TextBox txtVenta 
         Height          =   345
         Left            =   1725
         TabIndex        =   23
         Top             =   3270
         Width           =   2430
      End
      Begin VB.TextBox txtSubRubro 
         Height          =   345
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2610
         Width           =   555
      End
      Begin VB.TextBox txtRubro 
         Height          =   345
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1905
         Width           =   2430
      End
      Begin VB.TextBox txtProv 
         Height          =   345
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1320
         Width           =   2430
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "% Ocupacion"
         Height          =   315
         Left            =   210
         TabIndex        =   19
         Top             =   3930
         Width           =   1455
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sub Rubro"
         Height          =   315
         Left            =   210
         TabIndex        =   18
         Top             =   2625
         Width           =   1455
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Peso Venta"
         Height          =   315
         Left            =   210
         TabIndex        =   17
         Top             =   3285
         Width           =   1455
      End
      Begin VB.Label lblSectorProv 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   255
         TabIndex        =   16
         Top             =   630
         Width           =   3645
      End
      Begin VB.Label lblLocalProv 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   240
         TabIndex        =   15
         Top             =   150
         Width           =   3645
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor"
         Height          =   315
         Left            =   195
         TabIndex        =   14
         Top             =   1335
         Width           =   1455
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rubro"
         Height          =   315
         Left            =   210
         TabIndex        =   13
         Top             =   1905
         Width           =   1455
      End
   End
   Begin VB.Frame frSector 
      Height          =   5715
      Left            =   5460
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   4365
      Begin VB.CommandButton botDelSector 
         Caption         =   "Eliminar Sector"
         Height          =   465
         Left            =   2985
         TabIndex        =   30
         Top             =   5160
         Width           =   1290
      End
      Begin VB.TextBox txtCodSec 
         Height          =   285
         Left            =   1815
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1590
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton botOcupacion 
         Caption         =   "Modificación % de Ocupación"
         Height          =   465
         Left            =   1545
         TabIndex        =   26
         Top             =   5160
         Width           =   1290
      End
      Begin VB.CommandButton botNewProv 
         Caption         =   "Adicionar Proveedor"
         Height          =   465
         Left            =   120
         TabIndex        =   25
         Top             =   5160
         Width           =   1290
      End
      Begin VB.TextBox txtSup 
         Height          =   345
         Left            =   1740
         TabIndex        =   6
         Top             =   1545
         Width           =   2430
      End
      Begin VB.TextBox txtSector 
         Height          =   345
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1095
         Width           =   2430
      End
      Begin ComctlLib.ListView lstProv 
         Height          =   3075
         Left            =   225
         TabIndex        =   11
         Top             =   2010
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   5424
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   327680
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         MouseIcon       =   "frmAdmSectorProveedor.frx":001C
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Sector"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Superf."
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblLocalSec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   255
         TabIndex        =   7
         Top             =   165
         Width           =   3645
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Superficie"
         Height          =   315
         Left            =   225
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sector"
         Height          =   315
         Left            =   225
         TabIndex        =   3
         Top             =   1110
         Width           =   1455
      End
      Begin VB.Label lblSector 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   270
         TabIndex        =   2
         Top             =   660
         Width           =   3645
      End
   End
   Begin VB.Frame frLocal 
      Height          =   5715
      Left            =   5460
      TabIndex        =   8
      Top             =   30
      Visible         =   0   'False
      Width           =   4380
      Begin VB.CommandButton botSuperficie 
         Caption         =   "Modificación  de Superficie"
         Height          =   465
         Left            =   1770
         TabIndex        =   29
         Top             =   5145
         Width           =   1350
      End
      Begin VB.CommandButton botNewSector 
         Caption         =   "Adicionar Sector"
         Height          =   465
         Left            =   300
         TabIndex        =   28
         Top             =   5145
         Width           =   1350
      End
      Begin ComctlLib.ListView lstLocalSector 
         Height          =   4095
         Left            =   255
         TabIndex        =   9
         Top             =   990
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   7223
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   327680
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         MouseIcon       =   "frmAdmSectorProveedor.frx":0038
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Sector"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Superf."
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblLocal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   225
         TabIndex        =   10
         Top             =   285
         Width           =   3645
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   105
      Top             =   2430
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAdmSectorProveedor.frx":0054
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAdmSectorProveedor.frx":036E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdmSectorProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim moDragNode As Object ' Item that is being dragged.
Dim mbIndrag As Boolean

Public sqlProv As String
Public sqlSector As String

Dim Newsec As Integer
Private Sub ArmarArbol()
Dim n As Node

    TreeView1.Nodes.Clear
    'TreeView1.Nodes.Add TreeView1.SelectedItem.Index, tvwLast, skey, "Interbaires", 1, 2
    Set n = TreeView1.Nodes.Add(, tvwLast, "EMP1", "InterBaires", 1, 2)
    n.Tag = "EMPRESA"
    
    'Ingreso de las dependencias
    Set n = TreeView1.Nodes.Add(1, tvwChild, "INTA", "Internacional A ", 1, 2)
    n.Tag = "DEPN"
    L_LlenarLocalles "INTA", n.Index
    
    Set n = TreeView1.Nodes.Add(1, tvwChild, "INTB", "Internacional B ", 1, 2)
    n.Tag = "DEPN"
    L_LlenarLocalles "INTB", n.Index
    
    Set n = TreeView1.Nodes.Add(1, tvwChild, "AEP", "Aeroparque", 1, 2)
    n.Tag = "DEPN"
    L_LlenarLocalles "AEP", n.Index

    TreeView1.Nodes(1).Selected = True

    lstLocalSector.View = lvwReport
    lstProv.View = lvwReport

    Newsec = -1

End Sub

Public Sub EliminarNodo(Clave As String, indice As Integer)

If indice >= 0 Then
    TreeView1.Nodes.Remove (indice)
Else

End If

End Sub

Private Function GetNextKey() As String
'Returns a new key value for each Node being added to the TreeView
'This algorithm is very simple and will limit you to adding a total of 999 nodes
'Each node needs a unique key. If you allow users to remove Nodes you can't use
'the Nodes count +1 as the key for a new node.

    Dim sNewKey As String
    Dim iHold As Integer
    Dim i As Integer
    On Error GoTo myerr
    'The next line will return error #35600 if there are no Nodes in the TreeView
    iHold = Val(TreeView1.Nodes(1).Key)
    For i = 1 To TreeView1.Nodes.Count
        If Val(TreeView1.Nodes(i).Key) > iHold Then
            iHold = Val(TreeView1.Nodes(i).Key)
        End If
    Next
    iHold = iHold + 1
    sNewKey = CStr(iHold) & "_"
    GetNextKey = sNewKey 'Return a unique key
    Exit Function
myerr:
    'Because the TreeView is empty return a 1 for the key of the first Node
    GetNextKey = "1_"
    Exit Function
End Function


Private Sub L_LlenarListaLocalSector()
Dim mItem As ListItem

Dim rs As Recordset

lstLocalSector.ColumnHeaders.Clear
lstLocalSector.ColumnHeaders.Add , , "Nombre sector", 2000
lstLocalSector.ColumnHeaders.Add , , "Superf", 600

 sqlSector = "Select s.cod_sector,descrip, metros_cuadrados m2 "
 sqlSector = sqlSector & " FROM ESTADIS.Z_SECTOR s,estadis.z_sector_local  l "
 sqlSector = sqlSector & " Where s.cod_sector = l.cod_sector "
 sqlSector = sqlSector & " and cod_local = '" & TreeView1.SelectedItem.Key & "'"
 sqlSector = sqlSector & " order by descrip desc "
 
lstLocalSector.ListItems.Clear

If Aplicacion.ObtenerRsDAO(sqlSector, rs) Then
    Do While Not rs.EOF
        Set mItem = lstLocalSector.ListItems.Add(1, , rs!Descrip)
        mItem.SubItems(1) = rs!m2
    rs.MoveNext
    Loop
End If

'lstLocalSector.ListItems.Add 2, , "200 "

End Sub
Private Sub L_LlenarListaSectorProv()
Dim mItem As ListItem
Dim sql As String
Dim rs As Recordset

lstProv.ColumnHeaders.Clear
lstProv.ColumnHeaders.Add , , "Proveedor", 1500
lstProv.ColumnHeaders.Add , , "Rubro", 600
lstProv.ColumnHeaders.Add , , "Ocupa%", 600
 
 
 txtSector.Text = TreeView1.SelectedItem.Text
 txtCodSec.Text = Mid(TreeView1.SelectedItem.Key, InStr(1, TreeView1.SelectedItem.Key, "*") + 1)
 
 sql = "Select metros_cuadrados mc from estadis.z_sector_local where cod_local = '" & TreeView1.SelectedItem.Parent.Key & "' and cod_sector = " & txtCodSec.Text
 If Aplicacion.ObtenerRsDAO(sql, rs) Then
    If Not rs.EOF Then
        txtSup.Text = rs!mc & "  mtrs²"
    End If
    rs.Close
 End If
 
 sqlProv = " SELECT p.cod_prov,descrip, COD_RUBR, PARTICIPACION_OCUPA oc,secuencia "
 sqlProv = sqlProv & " FROM ESTADIS.Z_SECTOR_LOCAL_PROV z, baires.proveedor p "
 sqlProv = sqlProv & " Where p.cod_prov=z.cod_prov "
 sqlProv = sqlProv & " and cod_local = '" & TreeView1.SelectedItem.Parent.Key & "'"
 sqlProv = sqlProv & " and cod_sector = " & txtCodSec.Text
 sqlProv = sqlProv & " order by descrip desc "
 
 
lstProv.ListItems.Clear
If Aplicacion.ObtenerRsDAO(sqlProv, rs) Then
    Do While Not rs.EOF
        Set mItem = lstProv.ListItems.Add(1, , rs!Descrip)
        mItem.SubItems(1) = rs!cod_rubr
        mItem.SubItems(2) = rs!oc
    rs.MoveNext
    Loop
    rs.Close
End If


'lstLocalSector.ListItems.Add 2, , "200 "

End Sub

Private Sub L_LlenarLocalles(Dep As String, ent As Integer)
Dim n As Node
Dim sql As String
Dim rs As Recordset
    

sql = "Select cod_loc from baires.local where tipo_loc <> 'X' and cod_sdep = '" & Dep & "' "

If Aplicacion.ObtenerRsDAO(sql, rs) Then
    Do While Not rs.EOF
        Set n = TreeView1.Nodes.Add(ent, tvwChild, rs!cod_loc, "Local " & rs!cod_loc, 1, 2)
        n.Tag = "LOCAL"
            L_CargarLocalSector n
        rs.MoveNext
    Loop
End If

'  Set n = TreeView1.Nodes.Add(ent, tvwChild, "L06", "Local 06", 1, 2)
'  n.Tag = "LOCAL"
  

End Sub

Private Sub L_CargarLocalSector(n As Node)
Dim rs As Recordset
Dim sql As String
Dim nodo As Node

 sql = " Select s.cod_sector, descrip, metros_cuadrados m2 "
 sql = sql & " FROM ESTADIS.Z_SECTOR s,estadis.z_sector_local  l "
 sql = sql & " Where s.cod_sector = l.cod_sector "
 sql = sql & " and cod_local = '" & n.Key & "'"
 sql = sql & " order by descrip "

If Aplicacion.ObtenerRsDAO(sql, rs) Then
    Do While Not rs.EOF
        Set nodo = TreeView1.Nodes.Add(n.Index, tvwChild, n.Key & "*" & rs!cod_sector, "" & rs!Descrip, 1, 2)
        nodo.Tag = "SECTOR"

        L_CargarProveedor nodo
    rs.MoveNext
    Loop
End If

End Sub

Private Sub L_CargarProveedor(n As Node)
Dim sql As String
Dim rs As Recordset
Dim nodo As Node

 sql = " SELECT p.cod_prov,descrip, COD_RUBR, PARTICIPACION_OCUPA, secuencia "
 sql = sql & " FROM ESTADIS.Z_SECTOR_LOCAL_PROV z, baires.proveedor p "
 sql = sql & " Where p.cod_prov=z.cod_prov "
 sql = sql & " and cod_local = '" & n.Parent.Key & "'"
 sql = sql & " and cod_sector = " & Mid(n.Key, InStr(1, n.Key, "*") + 1)
 sql = sql & " order by descrip  "
 
If Aplicacion.ObtenerRsDAO(sql, rs) Then
    Do While Not rs.EOF
        Set nodo = TreeView1.Nodes.Add(n.Index, tvwChild, n.Key & "#" & rs!cod_prov & "@" & rs!secuencia, "" & rs!Descrip, 1, 2)
        nodo.Tag = "PROV"

    rs.MoveNext
    Loop
End If

End Sub


Private Sub L_LllenarPantProveedor(n As Node)
Dim sql As String
Dim rs As Recordset
Dim desc As String

 sql = " SELECT descrip, COD_RUBR,nvl(cod_srub,' ') cod_srub, PARTICIPACION_OCUPA oc,participacion_venta pv "
 sql = sql & " FROM ESTADIS.Z_SECTOR_LOCAL_PROV z, baires.proveedor p "
 sql = sql & " Where p.cod_prov=z.cod_prov "
 sql = sql & " and cod_local = '" & n.Parent.Parent.Key & "'"
 sql = sql & " and cod_sector = " & Mid(n.Parent.Key, InStr(1, n.Parent.Key, "*") + 1)
 sql = sql & " and z.cod_prov = '" & Mid(n.Key, InStr(1, n.Key, "#") + 1, InStr(1, n.Key, "@") - InStr(1, n.Key, "#") - 1) & "' "
 sql = sql & " and secuencia = " & Mid(n.Key, InStr(1, n.Key, "@") + 1)
 
 sql = sql & " order by descrip desc "

If Aplicacion.ObtenerRsDAO(sql, rs) Then
    Do While Not rs.EOF
        txtProv.Text = rs!Descrip
        txtRubro.Text = rs!cod_rubr
        txtSubRubro.Text = rs!cod_srub
        txtPart.Text = rs!oc
        txtVenta.Text = rs!pv
    rs.MoveNext
    Loop
End If

If Not IsNull(txtSubRubro.Text) Then
    sql = "Select descr descrip from baires.subrubro Where cod_rubr = '" & txtRubro.Text & "' "
    sql = sql & " And cod_srub = '" & txtSubRubro.Text & "' "
    
    Call Func.Func_ObtenerDesc(sql, desc)
    txtDescripSR.Text = desc
End If

End Sub

Private Sub botArbol_Click()
ArmarArbol
End Sub

Private Sub botDelProv_Click()
Dim sql As String
Dim n As Node

Set n = TreeView1.SelectedItem

If MsgBox("Confirma eliminación del proveedor " & txtProv.Text & " para el " & lblSectorProv.caption, vbYesNo + vbExclamation, "Atención") = vbYes Then
    sql = " Delete From estadis.z_sector_local_prov Where cod_local = '" & n.Parent.Parent.Key & "' "
    sql = sql & " And cod_sector = " & Mid(n.Parent.Key, InStr(1, n.Parent.Key, "*") + 1)
    sql = sql & " and cod_prov = '" & Mid(n.Key, InStr(1, n.Key, "#") + 1, InStr(1, n.Key, "@") - InStr(1, n.Key, "#") - 1) & "' "
    sql = sql & " and secuencia = " & Mid(n.Key, InStr(1, n.Key, "@") + 1)
    
    
    Aplicacion.ComienzoTrans
    
    If Aplicacion.EjecutarDAO(sql) Then
        Aplicacion.TerminarConExitoTrans
        EliminarNodo "", TreeView1.SelectedItem.Index
    Else
        Aplicacion.TerminarConErrorTrans
    End If
    
End If


End Sub

Private Sub botDelSector_Click()
Dim sql As String

If MsgBox("Confirma eliminación del sector " & lblSector.caption & " para el " & lblLocalSec.caption, vbYesNo + vbExclamation, "Atención") = vbYes Then
    sql = " Delete From estadis.z_sector_local Where cod_local = '" & TreeView1.SelectedItem.Parent.Key & "' "
    sql = sql & " And cod_sector = " & txtCodSec.Text
    
    Aplicacion.ComienzoTrans
    
    If Aplicacion.EjecutarDAO(sql) Then
        Aplicacion.TerminarConExitoTrans
        EliminarNodo "", TreeView1.SelectedItem.Index
    Else
        Aplicacion.TerminarConErrorTrans
    End If
    
End If

End Sub


Private Sub botModProv_Click()
Dim nnSector As Integer
Dim nnSecuen As Integer

frmADMSectorProv.modificacion TreeView1.SelectedItem, Newsec, nnSector, nnSecuen

If nnSector <> -1 Then
    TreeView1.SelectedItem.Key = Left(TreeView1.SelectedItem.Key, InStr(1, TreeView1.SelectedItem.Key, "*")) & nnSector & Mid(TreeView1.SelectedItem.Key, InStr(1, TreeView1.SelectedItem.Key, "#"), InStr(1, TreeView1.SelectedItem.Key, "@") - InStr(1, TreeView1.SelectedItem.Key, "#") + 1) & nnSecuen
End If
Newsec = nnSector

End Sub

Private Sub botNewProv_Click()
Dim nnDescrip As String
Dim nnProv As String
Dim nnSec As Integer
Dim n As Node

frmADMSectorProv.altas TreeView1.SelectedItem, nnProv, nnDescrip, nnSec

If nnProv <> "*" Then
    Set n = TreeView1.Nodes.Add(TreeView1.SelectedItem.Index, tvwChild, TreeView1.SelectedItem.Key & "#" & nnProv & "@" & nnSec, nnDescrip, 1, 2)
    n.Tag = "PROV"
    
    L_LlenarListaSectorProv
End If

End Sub

Private Sub botNewSector_Click()
Dim nnDescrip As String
Dim nnSector As Integer
Dim n As Node

frmADMSectorLocal.altas TreeView1.SelectedItem.Key, nnSector, nnDescrip

If nnSector <> -1 Then
    Set n = TreeView1.Nodes.Add(TreeView1.SelectedItem.Index, tvwChild, TreeView1.SelectedItem.Key & "*" & nnSector, nnDescrip, 1, 2)
    n.Tag = "SECTOR"
    
    L_LlenarListaLocalSector
End If

End Sub

Private Sub botOcupacion_Click()
frmADMSectorOC.Show 1
L_LlenarListaSectorProv
End Sub

Private Sub botSuperficie_Click()
frmADMSectorSup.Show 1
L_LlenarListaLocalSector
End Sub

Private Sub Command1_Click()
Dim skey As String
Dim n As Node

    skey = GetNextKey()
    If TreeView1.SelectedItem.Tag = "SECTOR" Then
        Set n = TreeView1.Nodes.Add(TreeView1.SelectedItem.Index, tvwChild, skey, "Prov:" & skey, 1, 2)
        n.Tag = "PROV"
    ElseIf TreeView1.SelectedItem.Tag = "LOCAL" Then
        Set n = TreeView1.Nodes.Add(TreeView1.SelectedItem.Index, tvwChild, skey, "Sec:" & skey, 1, 2)
        n.Tag = "SECTOR"
    End If
End Sub


Private Sub Form_Load()
Dim skey As String
Dim n As Node
    
Me.Top = 500
Me.Height = 7000
Me.Left = 645
    On Error GoTo myerr 'if the treeview does not have a node selected
    
    skey = GetNextKey()
    
    ArmarArbol
    
myerr:
    'TreeView1.Nodes.Add , tvwLast, skey, "Last " & skey, 1, 2
    Exit Sub

End Sub


Private Sub TreeView1_DragDrop(Source As control, x As Single, y As Single)
Dim a
        
If TreeView1.DropHighlight Is Nothing Then
        mbIndrag = False
        Exit Sub
ElseIf TreeView1.DropHighlight.Tag = "SECTOR" Then
    If frProv.Visible Then
        Newsec = Mid(TreeView1.DropHighlight.Key, InStr(1, TreeView1.DropHighlight.Key, "*") + 1)
        Call botModProv_Click
        
        If Newsec <> -1 Then
        ' Set dragged node's parent property to the target node.
            On Error GoTo checkerror ' To prevent circular errors.
            Set moDragNode.Parent = TreeView1.DropHighlight
            Cls
            'Print TreeView1.DropHighlight.Text " is parent of " & moDragNode.Text
        End If
        
        Set TreeView1.DropHighlight = Nothing
        mbIndrag = False
        Set moDragNode = Nothing
        Exit Sub ' Exit if no errors occured.
    End If
End If

checkerror:

End Sub

Private Sub TreeView1_DragOver(Source As control, x As Single, y As Single, State As Integer)
Dim a
Dim xX As New StdFont

xX.Bold = True
xX.Name = "Arial"

    
    If mbIndrag = True Then
        ' Set DropHighlight to the mouse's coordinates.
          Set TreeView1.DropHighlight = TreeView1.HitTest(x, y)
          If Not TreeView1.DropHighlight Is Nothing Then
          If TreeView1.DropHighlight.Tag <> "SECTOR" Then
          End If
          End If
    End If
End Sub


Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
If Not TreeView1.HitTest(x, y) Is Nothing Then
    Set TreeView1.DropHighlight = TreeView1.HitTest(x, y)
    TreeView1.SelectedItem = TreeView1.HitTest(x, y)
    
    If TreeView1.HitTest(x, y).Tag = "PROV" Then
        Set moDragNode = TreeView1.SelectedItem ' Set the item being dragged.
    Else
        Set moDragNode = Nothing
    End If
    'Debug.Print TreeView1.HitTest(x, y).Text
    Set TreeView1.DropHighlight = Nothing

End If

End Sub


Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton And Not moDragNode Is Nothing Then
        mbIndrag = True
        
        TreeView1.DragIcon = TreeView1.SelectedItem.CreateDragImage
        TreeView1.Drag vbBeginDrag ' Drag operation.
    End If

End Sub


Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
If Node.Tag = "SECTOR" Then
    frSector.Visible = True
    lblSector.caption = TreeView1.SelectedItem.Text
    lblLocalSec.caption = TreeView1.SelectedItem.Parent.Text
    L_LlenarListaSectorProv
    frLocal.Visible = False
    frProv.Visible = False
ElseIf Node.Tag = "LOCAL" Then
    frLocal.Visible = True
    lblLocal.caption = TreeView1.SelectedItem.Text
    frSector.Visible = False
    frProv.Visible = False
    L_LlenarListaLocalSector
ElseIf Node.Tag = "PROV" Then
    frProv.Visible = True
    lblSectorProv.caption = TreeView1.SelectedItem.Parent.Text
    lblLocalProv.caption = TreeView1.SelectedItem.Parent.Parent.Text
    
    L_LllenarPantProveedor TreeView1.SelectedItem
    frLocal.Visible = False
    frSector.Visible = False

Else
    frLocal.Visible = False
    frSector.Visible = False
    frProv.Visible = False
End If

End Sub

