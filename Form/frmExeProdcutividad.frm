VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmExeProductividad 
   Caption         =   "Proceso de cálculo de productividad"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2595
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   4575
      Begin VB.CommandButton botSalvar 
         Caption         =   "Ejecutar"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2700
         TabIndex        =   1
         Top             =   1860
         Width           =   990
      End
      Begin MSMask.MaskEdBox mskFDesde 
         Height          =   300
         Left            =   2040
         TabIndex        =   2
         Top             =   630
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox mskFHasta 
         Height          =   300
         Left            =   2055
         TabIndex        =   3
         Top             =   1050
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   327680
         PromptInclude   =   0   'False
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
         Caption         =   "Fecha desde"
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
         Left            =   375
         TabIndex        =   5
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "fecha Hasta"
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
         Left            =   375
         TabIndex        =   4
         Top             =   1050
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmExeProductividad"
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



Private Sub botSalvar_Click()
Dim sql As String
Dim rs As Recordset
Dim i


sql = " Begin estadis.procesos_estadisticos_mes.ACTUALIZA_PRODUCTIVIDAD( " _
& func_ToDate(mskFDesde.FormattedText) & ", " _
& func_ToDate(mskFHasta.FormattedText) & " ); End ; "

frmExeProductividad.caption = Aplicacion.SeteoProceso(frmExeProductividad.caption)

Aplicacion.ComienzoTrans

If Aplicacion.EjecutarDAO(sql) Then
    Aplicacion.TerminarConExitoTrans
    MsgBox "Proceso terminado exitosamente ", vbOKOnly + vbExclamation, "Atención"
Else
    Aplicacion.TerminarConErrorTrans
    MsgBox "Proceso terminado con error ", vbOKOnly + vbCritical, "Error"
End If

frmExeProductividad.caption = Aplicacion.SeteoFin

End Sub

Private Sub Form_Load()
Dim sql As String
Dim rs As Recordset

Top = 1800
Left = 1800

mskFDesde.Text = Func.func_Dia1SegunMes_Anio(Month(Now), Year(Now))
mskFHasta.Text = Format(Now - 1, FTOFECHA)

End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub




Private Sub mskFDesde_LostFocus()
    
    If Not IsDate(mskFDesde.FormattedText) Then
        mskFDesde.Text = Func.func_Dia1SegunMes_Anio(Month(Now), Year(Now))
    End If
        
    mskFDesde.Text = Format$(mskFDesde.FormattedText, FTOFECHA)
    mskFHasta.Text = Format$(mskFDesde.FormattedText, FTOFECHA)
    

End Sub


Private Sub mskFHasta_LostFocus()
    If Not IsDate(mskFHasta.FormattedText) Then
        mskFHasta.Text = Format$(Now - 1, FTOFECHA)
    End If
    
    If CDate(mskFHasta.FormattedText) < CDate(mskFDesde.FormattedText) Then
        mskFHasta.Text = mskFDesde.Text
    End If
    
    mskFHasta.Text = Format$(mskFHasta.FormattedText, FTOFECHA)
    

End Sub


