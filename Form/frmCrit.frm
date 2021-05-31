VERSION 5.00
Begin VB.Form frmCrit 
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000018&
      Height          =   3450
      Left            =   90
      ScaleHeight     =   3390
      ScaleWidth      =   5295
      TabIndex        =   0
      Top             =   105
      Width           =   5355
      Begin VB.CommandButton bot 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4845
         TabIndex        =   12
         Top             =   60
         Width           =   285
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   8
         Top             =   1935
         Value           =   2  'Grayed
         Width           =   180
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   960
         Value           =   2  'Grayed
         Width           =   180
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Son descontados como Pasajero Atendido, aquellos Pasaportes asociados exclusivamente a Ticket en Cero o Devoluciones."
         Height          =   390
         Index           =   3
         Left            =   450
         TabIndex        =   11
         Top             =   2805
         Width           =   4545
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "LOCAL"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1005
         TabIndex        =   10
         Top             =   2355
         Width           =   645
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "A nivel                , un mismo número de Pasaporte en distintos locales, son considerados como Dos Pasajeros."
         Height          =   390
         Index           =   2
         Left            =   465
         TabIndex        =   9
         Top             =   2355
         Width           =   4470
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Espigón Cia.Aerea-Vuelo-Nacionalidad"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   2130
         TabIndex        =   7
         Top             =   2130
         Width           =   3030
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Siempre se  tienen en cuenta todos los números de Pasaportes distintos considerando "
         Height          =   390
         Index           =   1
         Left            =   465
         TabIndex        =   6
         Top             =   1920
         Width           =   4470
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cálculo de cantidad de Pasajeros Atendidos"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   45
         TabIndex        =   5
         Top             =   1485
         Width           =   5205
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Se tienen en cuenta todos los números de Tickets distintos cuyo importe sea mayor a Cero"
         Height          =   435
         Index           =   0
         Left            =   495
         TabIndex        =   3
         Top             =   930
         Width           =   4440
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cálculo de cantidad de Tickets"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   45
         TabIndex        =   2
         Top             =   525
         Width           =   5220
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Criterios adoptados"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   45
         TabIndex        =   1
         Top             =   15
         Width           =   5235
      End
   End
End
Attribute VB_Name = "frmCrit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bot_Click()
Unload Me
End Sub

Private Sub Form_Load()
Top = 1800
Left = 1800
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub


