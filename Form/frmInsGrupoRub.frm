VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInsGrupoRub 
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabGrRub 
      Height          =   2895
      Left            =   225
      TabIndex        =   1
      Top             =   630
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5106
      _Version        =   327680
      TabHeight       =   520
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Grupo ""A"""
      TabPicture(0)   =   "frmInsGrupoRub.frx":0000
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "sprGR(0)"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Grupo ""B"""
      TabPicture(1)   =   "frmInsGrupoRub.frx":001C
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "sprGR(1)"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "Grupo ""C"""
      TabPicture(2)   =   "frmInsGrupoRub.frx":0038
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "sprGR(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Begin FPSpread.vaSpread sprGR 
         Height          =   2280
         Index           =   2
         Left            =   -74895
         OleObjectBlob   =   "frmInsGrupoRub.frx":0054
         TabIndex        =   4
         Top             =   435
         Width           =   4560
      End
      Begin FPSpread.vaSpread sprGR 
         Height          =   2280
         Index           =   1
         Left            =   -74880
         OleObjectBlob   =   "frmInsGrupoRub.frx":03A2
         TabIndex        =   3
         Top             =   420
         Width           =   4560
      End
      Begin FPSpread.vaSpread sprGR 
         Height          =   2280
         Index           =   0
         Left            =   135
         OleObjectBlob   =   "frmInsGrupoRub.frx":06F0
         TabIndex        =   2
         Top             =   435
         Width           =   4560
      End
   End
   Begin VB.Label labEsp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   240
      TabIndex        =   0
      Top             =   105
      Width           =   4800
   End
End
Attribute VB_Name = "frmInsGrupoRub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fchD As String
Dim fchH As String

Dim rsEstim As Recordset
Private Sub L_Refrescar()
Dim sql As String

On Error GoTo ErrEstim:

sql = "Select rubro, "
sql = sql & " sum(porc) PORC"
sql = sql & " From estadis.porcentaje_rd "
sql = sql & " WHERE anio = " & Year(fchD)
sql = sql & " And Mes = " & Month(fchD)
sql = sql & " And Tipo_porc = 'I' "
sql = sql & " And Nivel = 0 "
sql = sql & " AND DIA BETWEEN " & func_ToDate(fchD) & " AND " & func_ToDate(fchH)
sql = sql & " GROUP BY RUBRO "

If Aplicacion.ObtenerRsDAO(sql, rsEstim) Then

    If Aplicacion.CantReg(rsEstim) > 0 Then
        L_LlenarGrillasEstimRub
    End If
 
   Aplicacion.CerrarDAO rsEstim

End If

ErrEstim:
    Exit Sub
End Sub

Public Sub MostrarGrupoRubro(Esp As String, FD As String, FH As String)

fchD = FD
fchH = FH
labEsp.caption = Esp

Me.Show

End Sub


Private Sub Form_Load()
Top = 2700
Left = 4200

L_Refrescar

End Sub


