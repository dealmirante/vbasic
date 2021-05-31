VERSION 5.00
Begin VB.Form FrmLogon 
   Caption         =   "                           Ingreso al Sistema"
   ClientHeight    =   3405
   ClientLeft      =   2295
   ClientTop       =   3105
   ClientWidth     =   4830
   ControlBox      =   0   'False
   Icon            =   "Frmlogon.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3405
   ScaleWidth      =   4830
   Begin VB.PictureBox SSFrame1 
      Height          =   3435
      Left            =   75
      ScaleHeight     =   3375
      ScaleWidth      =   4575
      TabIndex        =   4
      Top             =   -45
      Width           =   4635
      Begin VB.TextBox txtUsuario 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2130
         TabIndex        =   0
         Top             =   945
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2130
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1515
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Height          =   840
         Left            =   150
         TabIndex        =   5
         Top             =   2250
         Width           =   4305
         Begin VB.CommandButton cmdSalir 
            Cancel          =   -1  'True
            Caption         =   "&Salir"
            Height          =   420
            Left            =   2370
            TabIndex        =   3
            Top             =   270
            Width           =   1695
         End
         Begin VB.CommandButton cmdIngresar 
            Caption         =   "&Ingresar"
            Default         =   -1  'True
            Height          =   435
            Left            =   255
            TabIndex        =   2
            Top             =   270
            Width           =   1575
         End
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Conexión al Sistema"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   315
         TabIndex        =   8
         Top             =   105
         Width           =   3810
      End
      Begin VB.Label lUsuario 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Usuario :"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   690
         TabIndex        =   7
         Top             =   945
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contraseña :"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   660
         TabIndex        =   6
         Top             =   1515
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'indica si el usuario presiono ingresar o salir
Dim f_bRespuesta As Boolean





Private Sub cmdIngresar_Click()
  
  Aplicacion.username = UCase(txtUsuario.Text)
  Aplicacion.password = txtPassword.Text
 
  f_bRespuesta = True
 
 Unload Me
  
End Sub

Private Sub cmdSalir_Click()
 
 f_bRespuesta = False
 
 Unload Me

End Sub

Public Function PedirUsuario() As Boolean

'Mostrar formulario. Lo muestra como modal No deja hacer
'otra cosa que responder la pregunta


Me.Show 1

PedirUsuario = f_bRespuesta

End Function

Private Sub Form_Load()
txtUsuario.Text = GetSetting("Aplicaciones", "Estadis", "Usuario")
End Sub


