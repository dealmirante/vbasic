VERSION 5.00
Begin VB.Form frmDatosCajero 
   Caption         =   "Informacion del Cajero"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   3975
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6855
      Begin VB.TextBox TxtLegajo 
         Height          =   330
         Left            =   1770
         TabIndex        =   14
         Top             =   255
         Width           =   1635
      End
      Begin VB.PictureBox Picture1 
         Height          =   3600
         Left            =   3495
         ScaleHeight     =   3540
         ScaleWidth      =   3210
         TabIndex        =   11
         Top             =   285
         Width           =   3270
         Begin VB.Image imgFoto 
            Height          =   3015
            Left            =   90
            Stretch         =   -1  'True
            Top             =   270
            Width           =   3060
         End
      End
      Begin VB.TextBox txtDescNac 
         Height          =   345
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2370
         Width           =   1620
      End
      Begin VB.TextBox txtDescIng 
         Height          =   330
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3255
         Width           =   1620
      End
      Begin VB.TextBox txtDescDoc 
         Height          =   345
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2820
         Width           =   1620
      End
      Begin VB.TextBox txtNombre 
         Height          =   345
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1920
         Width           =   3285
      End
      Begin VB.TextBox txtApe 
         Height          =   345
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1095
         Width           =   3285
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   5430
         TabIndex        =   13
         Top             =   -45
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Legajo"
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
         Index           =   3
         Left            =   105
         TabIndex        =   12
         Top             =   255
         Width           =   1620
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F. de Nacimiento"
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
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   2385
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F. de Ingreso"
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
         Index           =   4
         Left            =   135
         TabIndex        =   9
         Top             =   3255
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro Documento"
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
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   2820
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1530
         Width           =   1620
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Apellido"
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
         Left            =   120
         TabIndex        =   2
         Top             =   705
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmDatosCajero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cajero As Integer


Public Sub DatosEmpleado(Caj As Variant)

Cajero = Caj

L_LlenarDatosEmpleado

End Sub


Private Sub L_LlenarCampos(RS As Recordset)
Dim NombreArchivo As String

On Error GoTo ErrFoto:

    TxtLegajo.Text = RS!Legajo
    txtApe.Text = RS!apellido
    txtNombre.Text = RS!nombre
    txtDescNac.Text = Format$(RS!fch_nacimiento, FTOFECHA)
    txtDescDoc.Text = RS!Nro_Doc
    txtDescIng.Text = Format$(RS!Fch_ingreso, FTOFECHA)

    NombreArchivo = RutaFotos & "P" & Trim$(RS!Legajo) & ".bmp"
    
    Label1.caption = Cajero
    
    If Dir(NombreArchivo) <> "" Then
        imgFoto.Picture = LoadPicture(NombreArchivo)
    End If


ErrFoto:
    Exit Sub
End Sub

Private Sub L_LlenarDatosEmpleado()
Dim sql As String
Dim RS As Recordset

On Error GoTo ErrCarga:

sql = "SELECT "
sql = sql & "apellido,"
sql = sql & "nombre,"
sql = sql & "emp.legajo,"
sql = sql & "fch_nacimiento,"
sql = sql & "fch_ingreso,"
sql = sql & "nro_doc "
sql = sql & "FROM personal.empleado Emp ,ventas.cajeros Caj "
sql = sql & "WHERE Emp.legajo = Caj.legajo "
sql = sql & " AND caj.cod_cajero = " & Cajero

If Aplicacion.ObtenerRsDAO(sql, RS) Then
    If Aplicacion.CantReg(RS) > 0 Then
        L_LlenarCampos RS
        Show 1
    Else
        MsgBox "No existen datos de la persona", vbOKOnly + vbInformation, "Atención"
    End If
    Aplicacion.CerrarDAO RS
End If

ErrCarga:
    Exit Sub
End Sub

Private Sub Form_Load()
Top = 1200
Left = 1200
End Sub

