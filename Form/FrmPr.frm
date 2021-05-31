VERSION 5.00
Begin VB.Form frmPrincipal 
   BorderStyle     =   0  'None
   ClientHeight    =   6405
   ClientLeft      =   300
   ClientTop       =   570
   ClientWidth     =   8700
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6405
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6090
      Left            =   195
      TabIndex        =   0
      Top             =   165
      Width           =   8340
      Begin VB.ListBox lstForms 
         Height          =   450
         Left            =   7050
         TabIndex        =   5
         Top             =   1980
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton mnusalir 
         Caption         =   "&S  A  L  I  R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   4
         Left            =   135
         TabIndex        =   4
         Top             =   1455
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.CommandButton mnusalir 
         Caption         =   "&Adm. de Claves personales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   1050
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.CommandButton mnusalir 
         Caption         =   "&Adm. de Personal"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   135
         TabIndex        =   2
         Top             =   660
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.CommandButton mnusalir 
         Caption         =   "&Sist. Evaluación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "resto de la información se esta procesando..."
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Index           =   4
         Left            =   4515
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "9,10,11,12  de 1997"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Index           =   3
         Left            =   4500
         TabIndex        =   9
         Top             =   975
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "9,10,11,12 (al corriente) de 1998"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Index           =   2
         Left            =   4515
         TabIndex        =   8
         Top             =   645
         Visible         =   0   'False
         Width           =   3240
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Actualizaciones completas para los meses:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Index           =   1
         Left            =   255
         TabIndex        =   7
         Top             =   660
         Visible         =   0   'False
         Width           =   4140
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "IMPORTANTE"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Index           =   0
         Left            =   255
         TabIndex        =   6
         Top             =   165
         Visible         =   0   'False
         Width           =   3675
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit




Private Sub Form_Load()
Dim P As Panel

Top = 90
Left = 300

'MeVerPass

End Sub





Private Sub mnusalir_Click(Index As Integer)

Select Case Index
    Case 0
    Case 2
    Case 4
        End
End Select
End Sub

