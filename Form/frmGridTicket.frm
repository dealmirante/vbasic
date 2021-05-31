VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmGridTicket 
   Caption         =   "Datos"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   9825
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   420
      Left            =   30
      TabIndex        =   2
      Top             =   15
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
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
            Key             =   ""
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
            Key             =   ""
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
            Key             =   ""
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
      Top             =   390
      Width           =   9660
      Begin FPSpread.vaSpread spr 
         Height          =   3165
         Left            =   90
         OleObjectBlob   =   "frmGridTicket.frx":0000
         TabIndex        =   1
         Top             =   225
         Width           =   9450
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
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   5
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmGridTicket.frx":03AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmGridTicket.frx":06C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmGridTicket.frx":09E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmGridTicket.frx":1214
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmGridTicket.frx":152E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmGridTicket"
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
spr.MaxRows = 0
Do While Not r_Rs.EOF
    spr.MaxRows = spr.MaxRows + 1
    spr.SetText 1, spr.MaxRows, str(r_Rs!Cod_prod)
    spr.SetText 2, spr.MaxRows, Trim(r_Rs!Descrip)
    spr.SetText 3, spr.MaxRows, Trim(r_Rs!cod_rubr)
    spr.SetText 4, spr.MaxRows, Trim(r_Rs!cantidad)
    spr.SetText 5, spr.MaxRows, Trim(r_Rs!Importe)
    spr.SetText 6, spr.MaxRows, Trim(r_Rs!promotor)
    r_Rs.MoveNext
Loop
'Spread.Spread_TotalesGrillas spr, 8, 8
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
Width = 9600
Height = 4600
 
'Spread.Spread_CargarGrilla r_Rs, sprFly
L_Llenargrilla

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim valor As Variant

Select Case Button.Key
    Case "A"
        If Spread_FilaOcupada(spr, spr.MaxRows) Then
           Spread_AddRow spr
        End If
    Case "B"
        Spread_DelOneRow spr, spr.ActiveRow
        'DatoValido = True
    Case "C"
        spr.MaxRows = 0
    Case "D"
        Resp = 0
        Unload Me
    Case "E"
        If spr.ActiveRow = spr.MaxRows Then
            Resp = spr.ActiveRow - 1
        Else
            Resp = spr.ActiveRow
        End If
        Unload Me
        
End Select

End Sub






