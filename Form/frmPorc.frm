VERSION 5.00
Begin VB.Form frmPorc 
   Caption         =   "Form1"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1200
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   3765
      Begin VB.TextBox txtPorc 
         Height          =   315
         Left            =   1305
         TabIndex        =   4
         Text            =   "100"
         Top             =   405
         Width           =   1185
      End
      Begin VB.CommandButton bot 
         Caption         =   "Cancelar"
         Height          =   315
         Index           =   1
         Left            =   2730
         TabIndex        =   2
         Top             =   675
         Width           =   840
      End
      Begin VB.CommandButton bot 
         Caption         =   "Aceptar"
         Height          =   315
         Index           =   0
         Left            =   2730
         TabIndex        =   1
         Top             =   225
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   420
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmPorc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim P As Integer
Public Function TomarPorc() As Single
    
Me.Show 1

TomarPorc = P

End Function


Private Sub bot_Click(Index As Integer)
If Index = 0 Then
    P = txtPorc.Text
Else
    P = -1
End If
Unload Me
End Sub


Private Sub txtPorc_LostFocus()

If txtPorc.Text = "" Then
    txtPorc.Text = 100
End If
If Val(txtPorc.Text) < 0 Or Val(txtPorc.Text) > 100 Then
   txtPorc.Text = 100
End If

End Sub

