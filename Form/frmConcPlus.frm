VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmConcPlus 
   Caption         =   "$ Plus por Unidades Vendidas"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   420
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   4395
      _ExtentX        =   7752
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
            ImageIndex      =   16
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
            ImageIndex      =   17
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
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "D"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame FraObjetivo 
      Height          =   3315
      Left            =   15
      TabIndex        =   1
      Top             =   390
      Width           =   4380
      Begin VB.Frame Frame1 
         Height          =   2865
         Left            =   45
         TabIndex        =   4
         Top             =   405
         Width           =   4290
         Begin FPSpread.vaSpread sprPMP 
            Height          =   2220
            Left            =   225
            OleObjectBlob   =   "frmConcPlus.frx":0000
            TabIndex        =   5
            Top             =   495
            Width           =   3900
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "por Diferencia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   555
            TabIndex        =   7
            Top             =   195
            Width           =   1605
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "por Tope"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   2415
            TabIndex        =   6
            Top             =   195
            Value           =   -1  'True
            Width           =   1605
         End
      End
      Begin VB.OptionButton OptObjetivo 
         Caption         =   "Unidades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   615
         TabIndex        =   3
         Top             =   195
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptObjetivo 
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2475
         TabIndex        =   2
         Top             =   180
         Width           =   1740
      End
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   4440
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":02C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":05E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":08FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":0C16
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":0F30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":124A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":1564
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":187E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":1B98
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":1EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":1FC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":20D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":2678
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":2EAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":3668
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":3F2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConcPlus.frx":4244
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConcPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim col_Prod As Collection

Dim ModoEdit As Integer
Dim DatoValido As Boolean
Private Sub AdicionarAColeccion(cl As CLProdPrec)
Dim cl_Prod As CLProdPrec
Dim Resp As Boolean

Resp = True

For Each cl_Prod In col_Prod
    If cl.codProd = cl_Prod.codProd Then
        Resp = False
        Exit For
    End If
Next

If Resp Then
    col_Prod.Add cl
End If

End Sub

Private Sub L_Llenargrilla()
Dim cl As CLProdPrec

sprPMP.MaxRows = 0
For Each cl In col_Prod
    
    sprPMP.MaxRows = sprPMP.MaxRows + 1
    
    sprPMP.SetText 1, sprPMP.MaxRows, cl.codProd
    sprPMP.SetText 2, sprPMP.MaxRows, cl.Precio
    
    optTipo(cl.Tipo).Value = True
    OptObjetivo(cl.Objetivo).Value = True
    
    If cl.Tipo = 1 Then
        sprPMP.col = 2
       sprPMP.ColHidden = True
    End If
Next

End Sub
Private Sub LlenarColeccion()
Dim cl_Prod As CLProdPrec
Dim valor As Variant
Dim i

Set col_Prod = New Collection

For i = 1 To sprPMP.MaxRows
    
    Set cl_Prod = New CLProdPrec
                    
    sprPMP.GetText 1, i, valor
    If valor > -1 Then
        cl_Prod.codProd = valor
        
        sprPMP.GetText 2, i, valor
        cl_Prod.Precio = valor
                
        If optTipo(0).Value Then
            cl_Prod.Tipo = 0
        Else
            cl_Prod.Tipo = 1
        End If
        
        If OptObjetivo(0).Value Then
            cl_Prod.Objetivo = 0
        Else
            cl_Prod.Objetivo = 1
        End If
        
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
Width = 4900
Height = 4000

'UTILIZA la clase prodPrec para los plus donde el codigo es el valor

DatoValido = True
L_Llenargrilla
'OptObjetivo(0).Value = True

End Sub

Private Sub OptObjetivo_Click(Index As Integer)
Dim i As Integer

If Index = 1 Then
    optTipo(1).Value = True
    optTipo(0).Enabled = False
    sprPMP.col = 2
    sprPMP.ColHidden = True
    For i = 1 To sprPMP.MaxRows
        sprPMP.SetText 2, i, "0"
    Next

Else
    optTipo(0).Enabled = True
End If

End Sub

Private Sub OptTipo_Click(Index As Integer)
Dim i As Integer

If Index = 1 Then
    sprPMP.col = 2
    sprPMP.ColHidden = True
    For i = 1 To sprPMP.MaxRows
        sprPMP.SetText 2, i, "0"
    Next
Else
    sprPMP.col = 2
    sprPMP.ColHidden = False

End If

End Sub

Private Sub sprPMP_EditMode(ByVal col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
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
'    sprPMP.GetText 1, sprPMP.ActiveRow, cod
'
'    sql = "SELECT descrip FROM producto where cod_prod = " & cod
'
'    DatoValido = Func_ObtenerDesc(sql, desc)
'
'    sprPMP.SetText 2, sprPMP.ActiveRow, desc
'
    DatoValido = True
    If sprPMP.ActiveRow = sprPMP.MaxRows Then
       Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
    End If

End If

End Sub
Private Sub sprPMP_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim sql As String
Dim desc As String
Dim cod As Variant

If ModoEdit = 1 Then
    
    desc = ""
    ModoEdit = 0
    'sprPMP.GetText 1, Row, cod
   '
   ' sql = "SELECT descrip FROM producto where cod_prod = " & cod
   '
   ' DatoValido = Func_ObtenerDesc(sql, desc)
   '
   ' sprPMP.SetText 2, Row, desc
   DatoValido = True
    
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






