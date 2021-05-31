VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmAdmPersonas 
   Caption         =   "Administración de "
   ClientHeight    =   4035
   ClientLeft      =   705
   ClientTop       =   1170
   ClientWidth     =   5940
   Icon            =   "FrmAdmPersonas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4035
   ScaleWidth      =   5940
   Begin ComctlLib.Toolbar Tollbar 
      Height          =   420
      Left            =   120
      TabIndex        =   7
      Top             =   15
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   17
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "k"
            Object.ToolTipText     =   "Nueva Seleción"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "l"
            Object.ToolTipText     =   "Buscar"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "a"
            Object.ToolTipText     =   "Primer Registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "b"
            Object.ToolTipText     =   "Registro Anterior"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "c"
            Object.ToolTipText     =   "Registro Siguiente"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "e"
            Object.ToolTipText     =   "Ultimo Registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "f"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "g"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "h"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "i"
            Object.ToolTipText     =   "Abortar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "m"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "n"
            Object.ToolTipText     =   "Vista General"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "j"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame frEq 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Items del Equipo"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   120
      TabIndex        =   13
      Top             =   1860
      Width           =   5685
      Begin VB.OptionButton optTime 
         Caption         =   "PART Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4140
         TabIndex        =   32
         Top             =   225
         Width           =   1350
      End
      Begin VB.OptionButton optTime 
         Caption         =   "FULL Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2385
         TabIndex        =   31
         Top             =   225
         Value           =   -1  'True
         Width           =   1350
      End
      Begin VB.Frame Frame1 
         Caption         =   "Turno"
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
         Height          =   720
         Index           =   1
         Left            =   210
         TabIndex        =   27
         Top             =   1185
         Width           =   2835
         Begin VB.OptionButton optGr 
            Caption         =   " 'C'"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   2025
            TabIndex        =   30
            Top             =   315
            Width           =   690
         End
         Begin VB.OptionButton optGr 
            Caption         =   " 'B'"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1110
            TabIndex        =   29
            Top             =   330
            Width           =   690
         End
         Begin VB.OptionButton optGr 
            Caption         =   " 'A'"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   28
            Top             =   330
            Width           =   690
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Espigón"
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
         Height          =   705
         Index           =   0
         Left            =   195
         TabIndex        =   23
         Top             =   420
         Width           =   2835
         Begin VB.OptionButton optEsp 
            Caption         =   "Aep"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   2070
            TabIndex        =   26
            Top             =   330
            Width           =   645
         End
         Begin VB.OptionButton optEsp 
            Caption         =   "Int 'B'"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1125
            TabIndex        =   25
            Top             =   330
            Width           =   915
         End
         Begin VB.OptionButton optEsp 
            Caption         =   "Int 'A'"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   24
            Top             =   330
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1425
         Left            =   3075
         TabIndex        =   14
         Top             =   480
         Width           =   2490
         Begin VB.Frame Frame2 
            Caption         =   "Rubros"
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
            Height          =   1455
            Left            =   90
            TabIndex        =   15
            Top             =   -30
            Width           =   2385
            Begin VB.CheckBox chkRub 
               Caption         =   "CAJ"
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
               Height          =   225
               Index           =   6
               Left            =   1650
               TabIndex        =   22
               Top             =   615
               Width           =   690
            End
            Begin VB.CheckBox chkRub 
               Caption         =   "ACC"
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
               Height          =   225
               Index           =   0
               Left            =   105
               TabIndex        =   21
               Top             =   285
               Width           =   690
            End
            Begin VB.CheckBox chkRub 
               Caption         =   "COS"
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
               Height          =   225
               Index           =   4
               Left            =   945
               TabIndex        =   20
               Top             =   630
               Width           =   690
            End
            Begin VB.CheckBox chkRub 
               Caption         =   "COM"
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
               Height          =   240
               Index           =   3
               Left            =   945
               TabIndex        =   19
               Top             =   285
               Width           =   765
            End
            Begin VB.CheckBox chkRub 
               Caption         =   "CIG"
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
               Height          =   225
               Index           =   2
               Left            =   105
               TabIndex        =   18
               Top             =   990
               Width           =   690
            End
            Begin VB.CheckBox chkRub 
               Caption         =   "PER"
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
               Height          =   225
               Index           =   5
               Left            =   945
               TabIndex        =   17
               Top             =   990
               Width           =   690
            End
            Begin VB.CheckBox chkRub 
               Caption         =   "BEB"
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
               Height          =   240
               Index           =   1
               Left            =   105
               TabIndex        =   16
               Top             =   630
               Width           =   690
            End
         End
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "chk"
      Height          =   195
      Left            =   5265
      TabIndex        =   8
      Top             =   300
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txtReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4425
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   450
      Width           =   465
   End
   Begin VB.TextBox txtCantReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5220
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   450
      Width           =   480
   End
   Begin VB.Frame frCab 
      Height          =   1395
      Left            =   105
      TabIndex        =   9
      Top             =   465
      Width           =   5670
      Begin VB.CommandButton botHelpEq 
         Height          =   285
         Left            =   5175
         Picture         =   "FrmAdmPersonas.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   810
         Width           =   375
      End
      Begin VB.TextBox txtDescEq 
         Height          =   285
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   810
         Width           =   2355
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   2565
      End
      Begin MSMask.MaskEdBox mskLegajo 
         Height          =   285
         Left            =   1620
         TabIndex        =   0
         Top             =   360
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskCod 
         Height          =   285
         Left            =   1620
         TabIndex        =   1
         Top             =   810
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.Label de 
         Caption         =   "de"
         Height          =   255
         Left            =   4875
         TabIndex        =   12
         Top             =   120
         Width           =   405
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cod. Equipo"
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
         Left            =   360
         TabIndex        =   11
         Top             =   810
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Legajo"
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
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   1185
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   -60
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPersonas.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPersonas.frx":085E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPersonas.frx":0B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPersonas.frx":0E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPersonas.frx":11AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPersonas.frx":14C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPersonas.frx":17E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPersonas.frx":1AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPersonas.frx":1E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPersonas.frx":212E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPersonas.frx":2240
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPersonas.frx":2352
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmPersonas.frx":28F4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAdmPersonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rs As Recordset

Dim cl_P As CLEquipo

Dim CondConsulta As String

Dim Modo As String

Dim sqlGral$
Dim HelpEq As Boolean


Public Sub Altas()
    SetearBotonesAltas
    Modo = "ALTAS"
    FrmAdmPersonas.caption = FrmAdmPersonas.caption & " -Altas- "
    Me.Show 1
End Sub
Private Sub L_AltasDatos()

If L_TodoCargado Then
FrmAdmPersonas.caption = Aplicacion.SeteoProceso(FrmAdmPersonas.caption)

Aplicacion.ComienzoTrans

MeLlenarObjeto

If cl_P.Insert_Persona() Then
    Aplicacion.TerminarConExitoTrans
    chk.Value = 0
   
    NuevaSeleccion
    
Else
    Aplicacion.TerminarConErrorTrans
End If


FrmAdmPersonas.caption = Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub L_PonerEsp(Esp As String)

Select Case Esp
    Case "INTA"
        optEsp(0).Value = True
    Case "INTB"
        optEsp(1).Value = True
    Case "AEP"
        optEsp(2).Value = True
End Select
End Sub

Private Sub L_PonerGrupo(GR As String)

optGr(Asc(GR) - 65).Value = True

End Sub

Private Sub L_SegunTime(t As String)

Select Case t
    Case FULL
        optTime(0).Value = True
    Case part
        optTime(1).Value = True
End Select

End Sub

Private Function L_SegunTipo()

If optTime(0).Value Then
    L_SegunTipo = FULL
Else
    L_SegunTipo = part
End If

End Function

Private Sub MeImpDatos()
Dim nom As String, NombreArchivo As String


'On Error GoTo ErrFoto:
'
'Aplicacion.SeteoProceso ("")
'
'    NombreArchivo = RutaFotos & "P" & txtLegajo.Text & ".bmp"
'    Nom = txtApe.Text
'
'    If Dir(NombreArchivo) <> "" Then
'        Image1.Picture = LoadPicture(NombreArchivo)
'        Printer.PaintPicture Image1, 8000, 2000, 2800, 2200
'    End If
'
'
'Printer.FontBold = True
'
'Printer.CurrentX = 10
'Printer.CurrentY = 10
'Printer.FontSize = 10
'
'Printer.Print "  "
'
'Printer.CurrentX = 1000
'Printer.CurrentY = 1000
'Printer.FontSize = 18
'
'Printer.Print txtApe.Text & ", " & txtNom.Text
'
'Printer.CurrentX = 1000
'Printer.CurrentY = 2000
'Printer.FontSize = 10
''Printer.FontBold = False
'Printer.Print "Legajo  : "
'
'Printer.FontBold = False

'Printer.CurrentX = 10
'Printer.CurrentY = 10
'Printer.FontSize = 10
'
'Printer.Print "  "
'
'Printer.CurrentX = 2000
'Printer.CurrentY = 2000
'Printer.FontSize = 10
'Printer.Print txtLegajo.Text
'
'Printer.EndDoc
'
'ErrFoto:
'    Aplicacion.SeteoFin
'    Exit Sub
        
End Sub

Public Sub Modificacion()
Modo = "MOD"
FrmAdmPersonas.caption = FrmAdmPersonas.caption & " -Modificacion y Bajas- "
Me.Show 1
End Sub
Private Sub NuevaSeleccion()
Dim i

If Modo = "MOD" Then
    SetBotonesGeneral False
    mskCod.Text = ""
    For i = 0 To 5
        chkRub(i).Value = 0
    Next
    For i = 0 To 2
        optGr(i).Value = False
        optEsp(i).Value = False
    Next
Else
    If chk.Value = 1 Then
        If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
            L_AltasDatos
        End If
    End If
End If
'Limpiar campos de pantallas
Set cl_P = New CLEquipo

mskLegajo.Text = ""
mskCod.Text = ""
txtDesc.Text = ""
txtDescEq.Text = ""
'if msklegajo.Enabled then
mskLegajo.SetFocus
chk.Value = 0

End Sub

Private Sub MeAbortarMod()
    
    SeteoBotonesMod True
    
    Tollbar.Buttons(2).Enabled = False
    
    MeSetearBotonesToolBar
    
    MellenarPantalla
    
End Sub

Private Sub MeActualizar()
Dim ViejoOrgan$
Dim Viejocargo%

If L_TodoCargado Then

FrmAdmPersonas.caption = Aplicacion.SeteoProceso(FrmAdmPersonas.caption)

Aplicacion.ComienzoTrans

MeLlenarObjeto


If cl_P.Update_Persona Then '
    Aplicacion.TerminarConExitoTrans
    SeteoBotonesMod True

    If MeReconsultar > 0 Then

    Tollbar.Buttons(2).Enabled = False

    MeSetearBotonesToolBar
    Else
            NuevaSeleccion
    End If

Else
    Aplicacion.TerminarConErrorTrans
End If


FrmAdmPersonas.caption = Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub

Private Sub MeCargarDatos()
Dim sql$

'frm_.caption = Aplicacion.SeteoProceso (frm_.caption)
        
CondConsulta = ArmarCondicion

sqlGral$ = "SELECT legajo,E.cod_equipo,time " _
& " FROM estadis.persona_equipos PE, estadis.equipos E " _
& " where E.cod_equipo = PE.cod_equipo " _
& CondConsulta & " ORDER BY COD_EQUIPO "


If Aplicacion.ObtenerRsDAO(sqlGral$, rs) Then
    txtCantReg.Text = Aplicacion.CantReg(rs)
    If txtCantReg.Text > 0 Then
        txtReg.Text = 1
        SetBotonesGeneral True
        MellenarPantalla
        MeSetearBotonesToolBar
    Else
        txtReg.Text = 0
    End If
End If

'frm_.caption = Aplicacion.Seteofin

End Sub
Private Sub MeEliminar()

If MsgBox("Esta seguro de eliminar el registro", vbYesNo + vbExclamation, "ATENCION") = vbYes Then

MeLlenarObjeto

FrmAdmPersonas.caption = Aplicacion.SeteoProceso(FrmAdmPersonas.caption)

Aplicacion.ComienzoTrans

If cl_P.Delete_Persona Then
    Aplicacion.TerminarConExitoTrans
    SeteoBotonesMod True

    If MeReconsultar > 0 Then

    Tollbar.Buttons(2).Enabled = False

    MeSetearBotonesToolBar
    Else
            NuevaSeleccion
    End If

Else
    Aplicacion.TerminarConErrorTrans
End If

FrmAdmPersonas.caption = Aplicacion.SeteoFin
End If

End Sub

Private Sub MeLlenarObjeto()
 
 cl_P.cod = mskCod.Text
 cl_P.Legajo = mskLegajo.Text
 cl_P.TipoTime = L_SegunTipo
 
End Sub

Private Function L_TodoCargado() As Boolean

If mskLegajo.Text = "" Or mskCod.Text = "" Then
    L_TodoCargado = False
Else
    L_TodoCargado = True
End If

End Function

Private Sub MellenarPantalla()

mskLegajo.Text = rs!Legajo
mskCod.Text = rs!cod_equipo

L_SegunTime rs!Time

Call mskCod_LostFocus
Call mskLegajo_LostFocus

End Sub

Private Sub SetBotonesGeneral(valor As Boolean)
    
    Tollbar.Buttons(1).Enabled = valor
    Tollbar.Buttons(2).Enabled = Not valor
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(5).Enabled = valor
    Tollbar.Buttons(6).Enabled = valor
    Tollbar.Buttons(7).Enabled = valor

    Tollbar.Buttons(9).Enabled = valor
    Tollbar.Buttons(10).Enabled = valor
'
    Tollbar.Buttons(16).Enabled = valor
'    TollBar.Buttons(13).Enabled = Not valor
'

'habilitar frames
frCab.Enabled = Not valor
    
    If Not valor Then
        txtReg.Text = 0
        txtCantReg.Text = 0
    End If

End Sub

Private Function ArmarCondicion()
Dim Con$

Con$ = ""
If mskLegajo.Text <> "" Then
    Con$ = Con$ & " And PE.legajo = " & mskLegajo.Text
End If

If mskCod.Text <> "" Then
    Con$ = Con$ & " And E.Cod_equipo = " & mskCod.Text
End If

If optGr(0).Value Then
    Con$ = Con$ & " And GRUPO = 'A' "
ElseIf optGr(1).Value Then
    Con$ = Con$ & " And GRUPO = 'B' "
ElseIf optGr(2).Value Then
    Con$ = Con$ & " And GRUPO = 'C' "
End If

If optEsp(0).Value Then
    Con$ = Con$ & " And COD_SDEP = 'INTA' "
ElseIf optEsp(1).Value Then
    Con$ = Con$ & " And COD_SDEP = 'INTB' "
ElseIf optEsp(2).Value Then
    Con$ = Con$ & " And COD_SDEP = 'AEP' "
End If

'If Con$ <> "" Then
'    Con$ = " WHERE " & Mid(Con$, 5, Len(Con$))
'End If

ArmarCondicion = Con$

End Function



Private Sub MePrepararMod()
    
    SeteoBotonesMod False

End Sub

Private Function MeReconsultar() As Integer
Dim sql$
Dim i%
    

If Aplicacion.ObtenerRsDAO(sqlGral$, rs) Then
        txtCantReg.Text = Aplicacion.CantReg(rs)
        If Val(txtReg.Text) > Val(txtCantReg.Text) Then
            txtReg.Text = txtCantReg.Text
        End If
        
        For i% = 1 To txtReg.Text - 1
            rs.MoveNext
        Next
        If txtCantReg.Text > 0 Then
            MellenarPantalla
        End If
        'MeSetearBotonesToolBar
        MeReconsultar = txtCantReg.Text
End If
End Function


Private Sub MeSetearBotonesToolBar()
Dim i%
Dim but As Button

If txtCantReg.Text = 0 Then
'    TollBar.Buttons(1).Enabled = False
'    TollBar.Buttons(2).Enabled = False
'    TollBar.Buttons(3).Enabled = False
'    TollBar.Buttons(4).Enabled = False
'    TollBar.Buttons(6).Enabled = False
'    TollBar.Buttons(7).Enabled = False
ElseIf txtCantReg.Text = 1 Then
    Tollbar.Buttons(4).Enabled = False
    Tollbar.Buttons(5).Enabled = False
    Tollbar.Buttons(6).Enabled = False
    Tollbar.Buttons(7).Enabled = False
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
ElseIf txtReg.Text = txtCantReg.Text Then
    Tollbar.Buttons(4).Enabled = True
    Tollbar.Buttons(5).Enabled = True
    Tollbar.Buttons(6).Enabled = False
    Tollbar.Buttons(7).Enabled = False
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
ElseIf txtReg.Text = 1 Then
    Tollbar.Buttons(4).Enabled = False
    Tollbar.Buttons(5).Enabled = False
    Tollbar.Buttons(6).Enabled = True
    Tollbar.Buttons(7).Enabled = True
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
Else
    Tollbar.Buttons(4).Enabled = True
    Tollbar.Buttons(5).Enabled = True
    Tollbar.Buttons(6).Enabled = True
    Tollbar.Buttons(7).Enabled = True
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
    
End If
    


End Sub



Private Sub SetearBotonesAltas()
    
    Tollbar.Buttons(1).Enabled = True
    Tollbar.Buttons(11).Enabled = True
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Visible = False
    Tollbar.Buttons(15).Visible = False
    
    Tollbar.Buttons(2).Visible = False
    
    Tollbar.Buttons(4).Visible = False
    Tollbar.Buttons(5).Visible = False
    Tollbar.Buttons(6).Visible = False
    Tollbar.Buttons(7).Visible = False
    
    Tollbar.Buttons(9).Visible = False
    Tollbar.Buttons(10).Visible = False

    Tollbar.Buttons(15).Visible = False
    Tollbar.Buttons(16).Visible = False
    
    txtCantReg.Visible = False
    txtReg.Visible = False
    de.Visible = False
End Sub

Private Sub SeteoBotonesMod(valor As Boolean)
    
    
    Tollbar.Buttons(1).Enabled = valor
    Tollbar.Buttons(2).Enabled = valor
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(5).Enabled = valor
    Tollbar.Buttons(6).Enabled = valor
    Tollbar.Buttons(7).Enabled = valor

    Tollbar.Buttons(9).Enabled = valor
    Tollbar.Buttons(10).Enabled = valor

    Tollbar.Buttons(12).Enabled = Not valor
    Tollbar.Buttons(13).Enabled = Not valor

    Tollbar.Buttons(15).Enabled = valor
    Tollbar.Buttons(16).Enabled = valor

'habilitar o des frames y/o campos
frCab.Enabled = Not valor
mskLegajo.Enabled = valor

End Sub





Private Sub botHelpEq_Click()
Dim cl As CLEquipo

Set cl = New CLEquipo

FrmAdmEquip.EquipoAyuda cl

HelpEq = True

If cl.cod <> 0 Then
    mskCod.Text = cl.cod
    txtDescEq.Text = cl.desc
    
    L_PonerRubros cl.Rubros
    L_PonerGrupo cl.Grupo
    L_PonerEsp cl.CodSdep
End If
If mskCod.Enabled Then
    mskCod.SetFocus
End If

End Sub

Private Sub L_PonerRubros(RUBR As String)
Dim rubT As String, rub As String
Dim pos, i

    For i = 0 To 6
        chkRub(i).Value = 0
    Next

rubT = RUBR

pos = InStr(1, rubT, "/")

Do While pos <> 0
    rub = Left$(rubT, pos - 1)
    For i = 0 To 6
        If chkRub(i).caption = rub Then
            chkRub(i).Value = 1
            Exit For
        End If
    Next
    rubT = Right(rubT, Len(rubT) - pos)
    pos = InStr(1, rubT, "/")
Loop
    For i = 0 To 6
        If chkRub(i).caption = rubT Then
            chkRub(i).Value = 1
            Exit For
        End If
    Next

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 4 Then
    Select Case KeyCode
        Case 78 'Nueva sel
            Call TollBar_ButtonClick(Tollbar.Buttons(1))
        Case 71 'Guardar
            If Tollbar.Buttons(12).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(12))
            End If
        Case 66 'Buscar
            If Modo = "MOD" Then
            Call TollBar_ButtonClick(Tollbar.Buttons(2))
            End If
        Case 83 'Salir
            Call TollBar_ButtonClick(Tollbar.Buttons(16))
    End Select
    If Modo = "MOD" And Val(txtCantReg.Text) > 0 Then
    Select Case KeyCode
        Case 37 'Izq
            If Tollbar.Buttons(5).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(5))
            End If
        Case 38 'Arriba
            Call TollBar_ButtonClick(Tollbar.Buttons(7))
        Case 40 'Abajo
            Call TollBar_ButtonClick(Tollbar.Buttons(4))
        Case 39 'Der
            If Tollbar.Buttons(6).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(6))
            End If
    End Select
    End If

End If
'Debug.Print KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub Form_Load()
Width = Screen.Width * 0.65
Height = Screen.Height * 0.6
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height - 500) / 2

Set cl_P = New CLEquipo

End Sub

Private Sub mskCod_LostFocus()
Dim sql As String
Dim rs As Recordset
Dim cod, i

    If Not HelpEq Then
        cod = IIf(mskCod.Text = "", -1, mskCod.Text)
        sql = " SELECT cod_equipo, " _
              & " descrip, " _
              & " cod_depn, " _
              & " cod_sdep, " _
              & " grupo, " _
              & " rubros " _
              & " FROM estadis.Equipos " _
              & " WHERE cod_equipo = " & cod
            
            If Aplicacion.ObtenerRsDAO(sql, rs) Then
                If Aplicacion.CantReg(rs) > 0 Then
                    txtDescEq.Text = IIf(IsNull(rs!Descrip), "", rs!Descrip)
                    L_PonerRubros rs!Rubros
                    L_PonerGrupo rs!Grupo
                    L_PonerEsp rs!cod_sdep
                Else
                    txtDescEq.Text = ""
                    L_PonerRubros ""
                    For i = 0 To 2
                        optGr(i).Value = False
                        optEsp(i).Value = False
                    Next
                End If
                Aplicacion.CerrarDAO rs
            End If
        
    End If

    HelpEq = False
    If mskLegajo.Enabled Then
        mskLegajo.SetFocus
    End If
End Sub


Private Sub mskLegajo_LostFocus()
Dim sql As String
Dim desc As String

If mskLegajo.Text <> "" Then
    sql = "SELECT Apellido || ', ' || Nombre as descrip FROM personal.empleado " _
    & " WHERE legajo = " & mskLegajo.Text
    
    If Func_ObtenerDesc(sql, desc) Then
        txtDesc.Text = desc
    Else
        txtDesc.Text = "Legajo no resgistrado"
    End If
End If

End Sub

Private Sub TollBar_ButtonClick(ByVal Button As ComctlLib.Button)
Dim a%
Dim pos As String
Dim saltear As Boolean

saltear = True

pos = txtReg.Text

Select Case Button.Key
    Case "a"
         saltear = False
         Func_MoverPrimero rs, pos
    Case "b"
         saltear = False
        Func_MoverAnterior rs, pos
    Case "c"
         saltear = False
        Func_MoverSiguiente rs, pos
    Case "e"
         saltear = False
        Func_MoverUltimo rs, pos
    Case "f"
         MePrepararMod
    Case "g"
         MeEliminar
    Case "h"
        If Modo = "MOD" Then
            MeActualizar
        Else
            L_AltasDatos
        End If
    Case "i"
        MeAbortarMod
    Case "j"
        If chk.Value = 1 Then
            If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
                If Modo = "MOD" Then
                    MeActualizar
                Else
                    L_AltasDatos
                End If
            End If
        End If
        Unload Me
    Case "k"
        NuevaSeleccion
    Case "l"
        MeCargarDatos
    Case "n"
        L_DatosGrilla
    
    Case "m"
        MeImpDatos
    
End Select

If Not saltear Then
    txtReg.Text = pos
    MellenarPantalla
    MeSetearBotonesToolBar
End If

End Sub
Private Sub L_DatosGrilla()
Dim i, sql As String
Dim nro As Integer

On Error GoTo DG:

sql = "SELECT PE.legajo,apellido || ', ' || nombre Ape" _
& " FROM estadis.persona_equipos PE, estadis.equipos E, personal.Empleado EM" _
& " where E.cod_equipo = PE.cod_equipo " _
& " And PE.legajo = EM.legajo " _
& CondConsulta & " ORDER BY PE.COD_EQUIPO "

nro = frmGridPerEquip.DatosGrilla(sql)

If nro > 0 Then
    rs.MoveFirst
    For i = 1 To nro - 1
        rs.MoveNext
    Next
    MellenarPantalla
    txtReg.Text = nro
    MeSetearBotonesToolBar

DG:
    Exit Sub
End If

End Sub


