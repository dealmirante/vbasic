VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmGridFly 
   Caption         =   "Datos"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   420
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   8
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "A"
            Object.ToolTipText     =   "Agreagar Fila"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "B"
            Object.ToolTipText     =   "Sacar Fila"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "C"
            Object.ToolTipText     =   "Limpiar Todo"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "E"
            Object.ToolTipText     =   "Buscar "
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "D"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   3525
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   9120
      Begin FPSpread.vaSpread sprFly 
         Height          =   3165
         Left            =   90
         OleObjectBlob   =   "frmGridFly.frx":0000
         TabIndex        =   1
         Top             =   225
         Width           =   8955
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
            NumListImages   =   5
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmGridFly.frx":0461
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmGridFly.frx":077B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmGridFly.frx":0A95
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmGridFly.frx":12C7
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmGridFly.frx":15E1
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmGridFly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim r_Rs As Recordset
Dim Resp As Integer
Private Sub L_Llenargrilla()
Dim sqlCia As String
Dim desc As String

'sprFly.MaxRows = 0
sprFly.MaxRows = 0
Do While Not r_Rs.EOF
    sprFly.MaxRows = sprFly.MaxRows + 1
    
    sprFly.SetText 1, sprFly.MaxRows, Format$(r_Rs!fch_vuelo, "dd-mm-yy")
    sprFly.SetText 2, sprFly.MaxRows, Trim(r_Rs!cod_depn)
    sprFly.SetText 3, sprFly.MaxRows, Trim(r_Rs!cod_sdep)
    
    sqlCia = "Select descrip from ventas.companias where cod_compania = " & r_Rs!cod_cia_aerea
    Call Func.Func_ObtenerDesc(sqlCia, desc)
    sprFly.SetText 4, sprFly.MaxRows, desc
    
    sprFly.SetText 5, sprFly.MaxRows, str(r_Rs!cod_vuelo)
    'sprFly.SetText 5, sprFly.MaxRows, Trim(r_Rs!cod_sdep)
    sprFly.SetText 6, sprFly.MaxRows, str(r_Rs!cn)
    sprFly.SetText 7, sprFly.MaxRows, str(r_Rs!ce)
    
    r_Rs.MoveNext
Loop
Spread.Spread_TotalesGrillas sprFly, 7, 3

End Sub

Public Function DatosGrilla(ByVal Sql As String) As Integer

If Aplicacion.ObtenerRsDAO(Sql, r_Rs) Then

    Me.Show 1

    DatosGrilla = Resp
Else
    DatosGrilla = 0
End If

End Function

Private Sub Form_Load()
Top = 1700
Left = 700
Width = 9100
Height = 4600

'Spread.Spread_CargarGrilla r_Rs, sprFly
L_Llenargrilla

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim valor As Variant

Select Case Button.Key
    Case "A"
        If Spread_FilaOcupada(sprFly, sprFly.MaxRows) Then
           Spread_AddRow sprFly
        End If
    Case "B"
        Spread_DelOneRow sprFly, sprFly.ActiveRow
        'DatoValido = True
    Case "C"
        sprFly.MaxRows = 0
    Case "D"
        Resp = 0
        Unload Me
    Case "E"
        If sprFly.ActiveRow = sprFly.MaxRows Then
            Resp = sprFly.ActiveRow - 1
        Else
            Resp = sprFly.ActiveRow
        End If
        Unload Me
        
End Select

End Sub






