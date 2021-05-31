VERSION 5.00
Begin VB.Form frmTM 
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TM 
      Interval        =   3000
      Left            =   285
      Top             =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   15
      TabIndex        =   0
      Top             =   -30
      Width           =   6930
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6225
         TabIndex        =   1
         Top             =   1035
         Width           =   555
      End
      Begin VB.Label labMsg 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   255
         TabIndex        =   2
         Top             =   345
         Width           =   6345
      End
   End
End
Attribute VB_Name = "frmTM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim activo As Integer
Private Sub Command1_Click()
activo = 0
Me.Hide
End Sub

Private Sub Form_Load()
Width = Screen.Width * 0.6
Height = Screen.Height * 0.2
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height - 500) / 2

activo = 0
'Me.Show

End Sub


Private Sub TM_Timer()
Dim Ruta As String
Dim control As String

On Error GoTo Lectura:

    'Ruta = "c:\claudio\mensaje.txt"
    If activo = 0 Then
        Ruta = App.Path & "\mensaje.txt"
        Open Ruta For Input Lock Read Write As #2
        Line Input #2, control
            
        labMsg.caption = control
        Me.Show
        activo = 1
        DoEvents
        Close #2
    End If
    Exit Sub
    
Lectura:
    labMsg.caption = ""
    Exit Sub
    
End Sub


