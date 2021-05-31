VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmProdMon 
   Caption         =   "Códigos de Productos"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   420
      Left            =   15
      TabIndex        =   1
      Top             =   0
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   741
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "A"
            Object.ToolTipText     =   "Agreagar Fila"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
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
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "D"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   2865
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   5145
      Begin FPSpread.vaSpread sprPMP 
         Height          =   2535
         Left            =   165
         OleObjectBlob   =   "frmProdMon.frx":0000
         TabIndex        =   2
         Top             =   210
         Width           =   4800
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   4710
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327680
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   4
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmProdMon.frx":0255
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmProdMon.frx":056F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmProdMon.frx":0889
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmProdMon.frx":10BB
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmProdMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim col_Prod As Collection

Dim ModoEdit As Integer
Dim DatoValido As Boolean
Private Sub AdicionarAColeccion(cl As CLProdMon)
Dim cl_Prod As CLProdMon
Dim Resp As Boolean

Resp = True

For Each cl_Prod In col_Prod
    If cl.Cod_prod = cl_Prod.Cod_prod Then
        Resp = False
        Exit For
    End If
Next

If Resp Then
    col_Prod.Add cl
End If

End Sub

Private Sub L_Llenargrilla()
Dim cl As CLProdMon

sprPMP.MaxRows = 0
For Each cl In col_Prod
    
    sprPMP.MaxRows = sprPMP.MaxRows + 1
    
    sprPMP.SetText 1, sprPMP.MaxRows, str(cl.Cod_prod)
    sprPMP.SetText 2, sprPMP.MaxRows, cl.Descrip
    
Next

End Sub

Private Sub LlenarColeccion()
Dim cl_Prod As CLProdMon
Dim valor As Variant
Dim i

Set col_Prod = New Collection

For i = 1 To sprPMP.MaxRows
    
    Set cl_Prod = New CLProdMon
                    
    sprPMP.GetText 1, i, valor
    If valor <> 0 Then
        cl_Prod.Cod_prod = valor
        
        sprPMP.GetText 2, i, valor
        cl_Prod.Descrip = valor
        
        AdicionarAColeccion cl_Prod
    End If
    
Next

End Sub

Public Sub ProductosMonitoreo(ByRef cl As Collection)

Set col_Prod = cl

Me.Show 1

Set cl = col_Prod

End Sub

Private Sub Form_Load()
Top = 1000
Left = 1600
Width = 5300
Height = 3800


L_Llenargrilla

End Sub

Private Sub sprPMP_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If Mode = 1 Then
    ModoEdit = Mode
End If
End Sub

Private Sub sprPMP_KeyPress(KeyAscii As Integer)
Dim sql As String
Dim desc As String
Dim cod As Variant
    
If KeyAscii = 13 And ModoEdit = 1 Then
    
    desc = ""
    ModoEdit = 0
    sprPMP.GetText 1, sprPMP.ActiveRow, cod
    
    sql = "SELECT descrip FROM producto where cod_prod = " & cod
    
    DatoValido = Func_ObtenerDesc(sql, desc)
    
    sprPMP.SetText 2, sprPMP.ActiveRow, desc
    
    If sprPMP.ActiveRow = sprPMP.MaxRows Then
       Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
    End If

End If

End Sub

Private Sub sprPMP_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim sql As String
Dim desc As String
Dim cod As Variant


If ModoEdit = 1 Then
    
    desc = ""
    ModoEdit = 0
    sprPMP.GetText 1, Row, cod
    
    sql = "SELECT descrip FROM producto where cod_prod = " & cod
    
    DatoValido = Func_ObtenerDesc(sql, desc)
    
    sprPMP.SetText 2, Row, desc
    
    
End If

Cancel = Not DatoValido

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim valor As Variant

Select Case Button.Key
    Case "A"
        If Spread_FilaOcupada(sprPMP, sprPMP.MaxRows) Then
           Spread_AddRow sprPMP
        End If
    Case "B"
        Spread_DelOneRow sprPMP, sprPMP.ActiveRow
        DatoValido = True
    Case "C"
        sprPMP.MaxRows = 0
    Case "D"
        LlenarColeccion
        Unload Me
        
End Select

End Sub






