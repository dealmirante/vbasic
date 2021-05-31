VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmGridPerEquip 
   Caption         =   "Datos"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   420
      Left            =   30
      TabIndex        =   2
      Top             =   15
      Width           =   5430
      _ExtentX        =   9578
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
      MouseIcon       =   "frmGridPerEquip.frx":0000
   End
   Begin VB.Frame Frame1 
      Height          =   3525
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   5460
      Begin FPSpread.vaSpread spr 
         Height          =   3165
         Left            =   90
         OleObjectBlob   =   "frmGridPerEquip.frx":001C
         TabIndex        =   1
         Top             =   225
         Width           =   5205
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
               Picture         =   "frmGridPerEquip.frx":031B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmGridPerEquip.frx":0635
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmGridPerEquip.frx":094F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmGridPerEquip.frx":1181
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmGridPerEquip.frx":149B
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmGridPerEquip"
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
    spr.SetText 1, spr.MaxRows, str(r_Rs!Legajo)
    spr.SetText 2, spr.MaxRows, Trim(r_Rs!ape)
    r_Rs.MoveNext
Loop
'Spread.Spread_TotalesGrillas spr, 8, 8
End Sub

Public Function DatosGrilla(ByVal sql As String) As Integer

If Aplicacion.ObtenerRsDAO(sql, r_Rs) Then

    Me.Show 1

    DatosGrilla = Resp
Else
    DatosGrilla = 0
End If

End Function

Private Sub Form_Load()
Top = 1700
Left = 700
Width = 8100
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
        'If spr.ActiveRow = spr.MaxRows Then
        '    Resp = spr.ActiveRow - 1
        'Else
            Resp = spr.ActiveRow
        'End If
        Unload Me
        
End Select

End Sub






