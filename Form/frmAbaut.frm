VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAbaut 
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
      BackColor       =   &H00C0C0C0&
      Height          =   3450
      Left            =   60
      ScaleHeight     =   3390
      ScaleWidth      =   5295
      TabIndex        =   0
      Top             =   105
      Width           =   5355
      Begin VB.CommandButton botSalvar 
         Caption         =   "Salvar"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3945
         TabIndex        =   10
         Top             =   2805
         Visible         =   0   'False
         Width           =   990
      End
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
         TabIndex        =   4
         Top             =   60
         Width           =   285
      End
      Begin MSMask.MaskEdBox mskFch 
         Height          =   300
         Index           =   0
         Left            =   2070
         TabIndex        =   7
         Top             =   1020
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFch 
         Height          =   300
         Index           =   1
         Left            =   2070
         TabIndex        =   8
         Top             =   1440
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFch 
         Height          =   300
         Index           =   2
         Left            =   2070
         TabIndex        =   9
         Top             =   1860
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFch 
         Height          =   300
         Index           =   3
         Left            =   2055
         TabIndex        =   11
         Top             =   2265
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Interior"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   390
         TabIndex        =   12
         Top             =   2265
         Width           =   1545
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Internacinal 'B'"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   405
         TabIndex        =   6
         Top             =   1860
         Width           =   1545
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Internacinal 'A'"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   405
         TabIndex        =   5
         Top             =   1440
         Width           =   1545
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aeroparque"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   405
         TabIndex        =   3
         Top             =   1020
         Width           =   1545
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actualización de Pasajeros Transitados"
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
         Caption         =   "Estado de la Información"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
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
Attribute VB_Name = "frmAbaut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fch(0 To 3) As String
Dim Sdep(0 To 3) As String

Dim Modo As String

Public Sub Muestra()
    Modo = "MU"
    Me.Show 1
End Sub
Public Sub Seteo()
    Modo = "SE"
    Me.Show 1
End Sub


Private Sub bot_Click()
Unload Me
End Sub

Private Sub botSalvar_Click()
Dim sql$
Dim RS As Recordset
Dim i


For i = 0 To 3
    sql$ = "UPDATE estadis.control_carga SET " _
    & " fch_actualizado = " & func_ToDate(MskFch(i).FormattedText) _
    & " WHERE cod_sdep = '" & Sdep(i) & "'"
    
    Aplicacion.ComienzoTrans
    
    If Aplicacion.EjecutarDAO(sql$) Then
        Aplicacion.TerminarConExitoTrans
    Else
        Aplicacion.TerminarConErrorTrans
    End If
Next

End Sub

Private Sub Form_Load()
Dim sql As String
Dim RS As Recordset

Top = 1800
Left = 1800

sql = "SELECT cod_sdep,fch_actualizado FROM estadis.control_carga "

If Aplicacion.ObtenerRsDAO(sql, RS) Then
    
    Do While Not RS.EOF
        Select Case RS!Cod_Sdep
            Case "AEP"
                fch(0) = Format$(RS!fch_actualizado, FTOFECHA)
            Case "INTA"
                fch(1) = Format$(RS!fch_actualizado, FTOFECHA)
            Case "INTB"
                fch(2) = Format$(RS!fch_actualizado, FTOFECHA)
            Case "INT"
                fch(3) = Format$(RS!fch_actualizado, FTOFECHA)
        End Select
        RS.MoveNext
    Loop
    
    Aplicacion.CerrarDAO RS
End If

Sdep(0) = "AEP"
Sdep(1) = "INTA"
Sdep(2) = "INTB"
Sdep(3) = "INT"

MskFch(0).Text = fch(0)
MskFch(1).Text = fch(1)
MskFch(2).Text = fch(2)
MskFch(3).Text = fch(3)

If Modo = "SE" Then
    botSalvar.Visible = True
    MskFch(0).Enabled = True
    MskFch(1).Enabled = True
    MskFch(2).Enabled = True
    MskFch(3).Enabled = True
End If

End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub


Private Sub mskFch_LostFocus(Index As Integer)
    
    If Not IsDate(MskFch(Index).FormattedText) Then
        MskFch(Index) = Format$(fch(Index), FTOFECHA)
    ElseIf CDate(fch(Index)) > CDate(MskFch(Index).FormattedText) Then
        MskFch(Index) = Format$(fch(Index), FTOFECHA)
    End If
    
    MskFch(Index).Text = Format$(MskFch(Index).FormattedText, FTOFECHA)

End Sub


