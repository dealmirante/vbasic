VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmControlAus 
   Caption         =   "Form1"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   3885
      Begin VB.CommandButton botEjecutar 
         Caption         =   "Cancelar"
         Height          =   360
         Index           =   1
         Left            =   1965
         TabIndex        =   5
         Top             =   1125
         Width           =   1275
      End
      Begin VB.CommandButton botEjecutar 
         Caption         =   "Aceptar"
         Height          =   360
         Index           =   0
         Left            =   315
         TabIndex        =   4
         Top             =   1140
         Width           =   1275
      End
      Begin VB.CommandButton botHelpFD 
         Height          =   345
         Left            =   2910
         Picture         =   "frmControlAus.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   435
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   285
         Left            =   1605
         TabIndex        =   2
         Top             =   435
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha "
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
         Left            =   315
         TabIndex        =   3
         Top             =   450
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmControlAus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub botEjecutar_Click(Index As Integer)
Dim sql As String

sql = " Begin personal.Resumen_Cuadro_Ausent(" & func_ToDate(mskFDesde.FormattedText) & ") ; End ;"

If Index = 0 Then
    frmControlAus.caption = Aplicacion.SeteoProceso(frmControlAus.caption)
    
    Aplicacion.ComienzoTrans
    
    If Aplicacion.EjecutarDAO(sql) Then
        Aplicacion.TerminarConExitoTrans
        MsgBox "Proceso de Actualizacion OK", vbOKOnly + vbExclamation, "Ausentismo"
    Else
        Aplicacion.TerminarConErrorTrans
        MsgBox "Proceso de Actualizacion Termino mal", vbOKOnly + vbCritical, "Ausentismo"
    End If
    
    frmControlAus.caption = Aplicacion.SeteoFin
Else
    Unload Me
End If

End Sub

Private Sub botHelpFD_Click()
Dim fch As Date

If mskFDesde.Text <> "" Then
    fch = mskFDesde.FormattedText
Else
    fch = Date
End If

frmFecha.MuestroFormFecha fch

mskFDesde.Text = Format$(fch, FTOFECHA)

mskFDesde.SetFocus

End Sub

Private Sub Form_Load()
Top = 4000
Left = 4000
End Sub

Private Sub mskFDesde_LostFocus()

'If mskFDesde.Text <> "" Then
    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Format$(Date - 1, FTOFECHA)
    End If
    
    mskFDesde.Text = Format$(mskFDesde.FormattedText, FTOFECHA)
'End If


End Sub


