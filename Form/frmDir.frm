VERSION 5.00
Begin VB.Form frmDir 
   Caption         =   "Nombre de Archivo Excel"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton botSiNo 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   4650
      TabIndex        =   7
      Top             =   3105
      Width           =   1275
   End
   Begin VB.CommandButton botSiNo 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   3240
      TabIndex        =   6
      Top             =   3105
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre Archivo (*.xls)"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   630
      Left            =   255
      TabIndex        =   4
      Top             =   15
      Width           =   5685
      Begin VB.TextBox txtNombre 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1395
         TabIndex        =   5
         Top             =   195
         Width           =   4200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2370
      Left            =   270
      TabIndex        =   0
      Top             =   645
      Width           =   5655
      Begin VB.FileListBox filArchivo 
         Height          =   1845
         Left            =   2955
         Pattern         =   "*.xls"
         TabIndex        =   3
         Top             =   345
         Width           =   2565
      End
      Begin VB.DirListBox dirArchivo 
         Height          =   1440
         Left            =   135
         TabIndex        =   2
         Top             =   330
         Width           =   2670
      End
      Begin VB.DriveListBox drvArchivo 
         Height          =   315
         Left            =   135
         TabIndex        =   1
         Top             =   1830
         Width           =   2670
      End
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nombre As String

Private Function L_SinPunto(texto As String)
Dim pos

pos = InStr(1, texto, ".")
If pos = 0 Then
    L_SinPunto = texto
Else
    L_SinPunto = Left(texto, pos - 1)
End If
End Function

Public Function NombreArchivo() As String
    
    Me.Show 1
    NombreArchivo = nombre
    
End Function
Public Function NombreArch(PAT As String) As String
    
    filArchivo.Pattern = "*." & PAT
    Me.Show 1
    If nombre <> "" Then
        NombreArch = L_SinPunto(nombre) & "." & PAT
    Else
        NombreArch = ""
    End If
    
End Function


Private Sub botSiNo_Click(Index As Integer)
Select Case Index
    Case 0
        If txtNombre.Text <> "" Then
            If Len(dirArchivo.Path) = 3 Then
                nombre = dirArchivo.Path & L_SinPunto(txtNombre.Text)
            Else
                nombre = dirArchivo.Path & "\" & L_SinPunto(txtNombre.Text)
            End If
            Unload Me
        End If
    Case 1
        nombre = ""
        Unload Me
End Select
End Sub

Private Sub dirArchivo_Change()
filArchivo.Path = dirArchivo.Path
txtNombre.Text = ""
End Sub

Private Sub drvArchivo_Change()
dirArchivo.Path = drvArchivo.Drive
txtNombre.Text = ""
End Sub


Private Sub filArchivo_Click()
    txtNombre.Text = filArchivo.List(filArchivo.ListIndex)
End Sub

Private Sub Form_Load()
Top = 1300
Left = 1200
Width = 6300
Height = 4000

On Error GoTo ErrDir:

'If Len(RutaVisual) > 3 Then
    drvArchivo.Drive = "C:"
    dirArchivo.Path = "c:\mis documentos\"
'End If
txtNombre.Locked = False
ErrDir:
    Exit Sub
End Sub


