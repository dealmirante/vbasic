VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMonitoreoMes 
   Caption         =   "Información anual"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   180
      TabIndex        =   2
      Top             =   5250
      Width           =   5130
   End
   Begin VB.Frame Frame2 
      Height          =   4155
      Left            =   165
      TabIndex        =   1
      Top             =   1020
      Width           =   5145
      Begin TabDlg.SSTab SSTab1 
         Height          =   3825
         Left            =   105
         TabIndex        =   6
         Top             =   180
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   6747
         _Version        =   327680
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   441
         TabCaption(0)   =   "IMPORTE"
         TabPicture(0)   =   "frmMonitoreoMes.frx":0000
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "grid(0)"
         Tab(0).Control(0).Enabled=   0   'False
         TabCaption(1)   =   "UNIDADES"
         TabPicture(1)   =   "frmMonitoreoMes.frx":001C
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "grid(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Begin FPSpread.vaSpread grid 
            Height          =   3420
            Index           =   1
            Left            =   -74955
            OleObjectBlob   =   "frmMonitoreoMes.frx":0038
            TabIndex        =   8
            Top             =   60
            Width           =   4800
         End
         Begin FPSpread.vaSpread grid 
            Height          =   3420
            Index           =   0
            Left            =   45
            OleObjectBlob   =   "frmMonitoreoMes.frx":0411
            TabIndex        =   7
            Top             =   60
            Width           =   4800
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   990
      Left            =   165
      TabIndex        =   0
      Top             =   -15
      Width           =   5145
      Begin VB.TextBox txtStk 
         Height          =   300
         Left            =   1860
         TabIndex        =   5
         Top             =   510
         Width           =   2850
      End
      Begin VB.Label lblProv 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   240
         Left            =   150
         TabIndex        =   4
         Top             =   30
         Width           =   4860
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stock "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   375
         TabIndex        =   3
         Top             =   510
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmMonitoreoMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim prov As String
Dim RUBR As String
Dim desc As String
Dim sql As String
Dim tipo As String

Dim RS As Recordset
Public Sub Mostrar(pSQL As String, pProv As String, pdescPv As String, pRubr As String, pTipo As String)

RUBR = pRubr
prov = pProv
desc = pdescPv
tipo = pTipo
sql = pSQL


Me.Show 1

End Sub

Private Sub L_Refrescar()
Dim sqlStk As String
Dim rsStk As Recordset
Dim vta As Variant, est As Variant

frmMonitoreoMes.caption = Aplicacion.SeteoProceso(frmMonitoreoMes.caption)

lblProv.caption = tipo & " : " & desc
DoEvents

If tipo = "Proveedor" Then
    If RUBR = "" Then
      sqlStk = "Select baires.Stock_prov('" & prov & "') stk From dual "
    Else
      sqlStk = "Select baires.Stk_x_prov_rubr('" & prov & "','" & RUBR & "' ) stk From dual "
    End If
Else
' prov hace las veces de producto
  sqlStk = "Select baires.Stock_cia('" & prov & "') stk From dual "
End If

If Aplicacion.ObtenerRsDAO(sqlStk, rsStk) Then
   If Aplicacion.CantReg(rsStk) > 0 Then
        txtStk.Text = rsStk!stk
   End If
   rsStk.Close
End If

If Aplicacion.ObtenerRsDAO(sql, RS) Then
    grid(0).MaxRows = 0
    grid(1).MaxRows = 0
    Do While Not RS.EOF
        
        grid(0).MaxRows = grid(0).MaxRows + 1
        grid(1).MaxRows = grid(1).MaxRows + 1
        
        grid(0).SetText 1, grid(0).MaxRows, str(RS!aniomes)
        grid(1).SetText 1, grid(1).MaxRows, str(RS!aniomes)

        grid(0).SetText 2, grid(0).MaxRows, str(RS!venta_real_P)
        grid(1).SetText 2, grid(1).MaxRows, str(RS!venta_real_U)

        grid(0).SetText 3, grid(0).MaxRows, str(RS!estimado_P)
        grid(1).SetText 3, grid(1).MaxRows, str(RS!estimado_U)

        RS.MoveNext
    Loop
    Spread.Spread_TotalesGrillas grid(0), 3, 2
    Spread.Spread_TotalesGrillas grid(1), 3, 2
    
    grid(0).GetText 2, grid(0).MaxRows, vta
    grid(0).GetText 3, grid(0).MaxRows, est
    If Val(est) > 0 Then
       grid(0).SetText 4, grid(0).MaxRows, ((vta - est) / est) * 100
    End If

    grid(1).GetText 2, grid(1).MaxRows, vta
    grid(1).GetText 3, grid(1).MaxRows, est
    If Val(est) > 0 Then
       grid(1).SetText 4, grid(1).MaxRows, ((vta - est) / est) * 100
    End If

End If

frmMonitoreoMes.caption = Aplicacion.SeteoFin

End Sub

Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Form_Load()
L_Refrescar

End Sub

