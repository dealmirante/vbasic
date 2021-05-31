VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form FrmAdmConcurso 
   Caption         =   "Administración de "
   ClientHeight    =   6105
   ClientLeft      =   705
   ClientTop       =   1170
   ClientWidth     =   7005
   Icon            =   "FrmAdmConcurso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6105
   ScaleWidth      =   7005
   Begin ComctlLib.Toolbar Tollbar 
      Height          =   420
      Left            =   120
      TabIndex        =   2
      Top             =   15
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   20
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "o"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "p"
            Object.ToolTipText     =   "Estado Mod / Consulta"
            Object.Tag             =   ""
            ImageIndex      =   14
            Value           =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "k"
            Object.ToolTipText     =   "Nueva Seleción"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "l"
            Object.ToolTipText     =   "Buscar"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "a"
            Object.ToolTipText     =   "Primer Registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "b"
            Object.ToolTipText     =   "Registro Anterior"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "c"
            Object.ToolTipText     =   "Registro Siguiente"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "e"
            Object.ToolTipText     =   "Ultimo Registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "f"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "g"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "h"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "i"
            Object.ToolTipText     =   "Abortar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "m"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "n"
            Object.ToolTipText     =   "Grilla"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "j"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   975
      Left            =   165
      TabIndex        =   25
      Top             =   2010
      Width           =   6585
      Begin VB.OptionButton optTipo 
         Caption         =   "Pago por unidades"
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
         Left            =   810
         TabIndex        =   31
         Top             =   180
         Width           =   1995
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Pago proporcional a importes"
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
         Left            =   3015
         TabIndex        =   30
         Top             =   180
         Width           =   2880
      End
      Begin MSMask.MaskEdBox mskP 
         Height          =   315
         Left            =   4770
         TabIndex        =   26
         Top             =   525
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   5
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskM 
         Height          =   315
         Left            =   1725
         TabIndex        =   27
         Top             =   525
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   5
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proporcional"
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
         Index           =   5
         Left            =   495
         TabIndex        =   29
         Top             =   525
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PAGA"
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
         Index           =   4
         Left            =   3540
         TabIndex        =   28
         Top             =   540
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3045
      Left            =   150
      TabIndex        =   6
      Top             =   2955
      Width           =   6585
      Begin FPSpread.vaSpread spr 
         Height          =   2265
         Left            =   105
         OleObjectBlob   =   "FrmAdmConcurso.frx":0442
         TabIndex        =   10
         Top             =   645
         Width           =   6345
      End
      Begin VB.Frame frBot 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   645
         Left            =   105
         TabIndex        =   7
         Top             =   195
         Width           =   6210
         Begin ComctlLib.Toolbar Toolbar1 
            Height          =   420
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   741
            ButtonWidth     =   635
            ButtonHeight    =   582
            Appearance      =   1
            ImageList       =   "ImageList1"
            _Version        =   327680
            BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
               NumButtons      =   7
               BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "A"
                  Object.ToolTipText     =   "Agreagar Fila"
                  Object.Tag             =   ""
                  ImageIndex      =   16
               EndProperty
               BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "B"
                  Object.ToolTipText     =   "Sacar Fila"
                  Object.Tag             =   ""
                  ImageIndex      =   17
               EndProperty
               BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "C"
                  Object.ToolTipText     =   "Limpiar Todo"
                  Object.Tag             =   ""
                  ImageIndex      =   9
               EndProperty
               BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Visible         =   0   'False
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Object.Visible         =   0   'False
                  Key             =   "D"
                  Object.ToolTipText     =   "Salir"
                  Object.Tag             =   ""
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin MSMask.MaskEdBox mskPaga 
            Height          =   285
            Left            =   4155
            TabIndex        =   14
            Top             =   60
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            _Version        =   327680
            PromptInclude   =   0   'False
            MaxLength       =   5
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   " "
         End
         Begin VB.CommandButton botHelpProd 
            Height          =   315
            Left            =   5565
            Picture         =   "FrmAdmConcurso.frx":076D
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Códigos de Productos"
            Top             =   60
            Width           =   405
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PAGA"
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
            Left            =   2835
            TabIndex        =   13
            Top             =   60
            Width           =   1185
         End
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "chk"
      Height          =   195
      Left            =   3045
      TabIndex        =   3
      Top             =   570
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txtReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4725
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   465
   End
   Begin VB.TextBox txtCantReg 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   480
   End
   Begin VB.Frame frCabecera 
      Height          =   1560
      Left            =   165
      TabIndex        =   4
      Top             =   450
      Width           =   6570
      Begin VB.Frame frCab 
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   135
         TabIndex        =   16
         Top             =   195
         Width           =   4815
         Begin MSMask.MaskEdBox msksec 
            Height          =   300
            Left            =   1350
            TabIndex        =   24
            Top             =   975
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            _Version        =   327680
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.ComboBox cboConcurso 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   120
            Width           =   2475
         End
         Begin VB.TextBox txtConc 
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   3855
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   19
            Top             =   120
            Width           =   1020
         End
         Begin VB.ComboBox cboProv 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   555
            Width           =   2475
         End
         Begin VB.TextBox txtProv 
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   3855
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   17
            Top             =   555
            Width           =   735
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Secuencia"
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
            Left            =   75
            TabIndex        =   23
            Top             =   975
            Width           =   1185
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Concurso"
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
            Left            =   75
            TabIndex        =   22
            Top             =   135
            Width           =   1185
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Proveedor"
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
            Left            =   75
            TabIndex        =   21
            Top             =   570
            Width           =   1185
         End
      End
      Begin VB.CommandButton botPlus 
         Height          =   945
         Left            =   5295
         Picture         =   "FrmAdmConcurso.frx":086F
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   435
         Width           =   1125
      End
      Begin VB.ListBox lstProv 
         Height          =   255
         Left            =   4920
         TabIndex        =   12
         Top             =   900
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ListBox LstConcurso 
         Height          =   255
         Left            =   4950
         TabIndex        =   11
         Top             =   630
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label de 
         Caption         =   "de"
         Height          =   255
         Left            =   5190
         TabIndex        =   5
         Top             =   135
         Width           =   315
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   165
      Top             =   255
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":1709
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":1A23
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":1D3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":2057
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":2371
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":268B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":29A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":2CBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":2FD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":32F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":3405
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":3517
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":3AB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":42EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":4AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":536B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAdmConcurso.frx":5685
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAdmConcurso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS As Recordset

Dim ModoEdit As Integer
Dim DatoValido As Boolean

Dim cl_ProvProd As CLConc
Dim col_Plus As Collection

Dim CondConsulta As String

Dim Modo As String

Dim sqlGral$
Public Sub Altas()
    SetearBotonesAltas True
    Modo = "ALTA"
    FrmAdmConcurso.caption = " Asignación de Productos -Altas- "
End Sub
Public Sub ConsultaBases(Conc As String)

SetearBotonesConsBases True

Modo = "BASE"

Me.Show
Me.caption = "Consultas de Bases del Concurso "
Func.Func_SetearCboConLst cboConcurso, LstConcurso, Conc

Call TollBar_ButtonClick(Tollbar.Buttons(5))

End Sub
Private Sub L_AltasDatos()

If L_TodoCargado Then
    FrmAdmConcurso.caption = Aplicacion.SeteoProceso(FrmAdmConcurso.caption)

    Aplicacion.ComienzoTrans

    MeLlenarObjeto

    If cl_ProvProd.Insert_Prod() Then
        Aplicacion.TerminarConExitoTrans
        chk.Value = 0
        
        NuevaSeleccion
        
    Else
        Aplicacion.TerminarConErrorTrans
    End If


    FrmAdmConcurso.caption = Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub
Private Sub L_LLenarPlus()
Dim sql As String
Dim rsP As Recordset
Dim cl_PP As CLProdPrec

sql = "SELECT limite_rango,plus ,tipo_pluss,tipo_objetivo" _
& " FROM estadis.concurso_plus " _
& " WHERE  id_concurso = '" & txtConc.Text & "' " _
& " and cod_prov = '" & txtProv.Text & "' " _
& " and secuencia = " & msksec.Text & " " _
& " ORDER BY limite_rango "

Set col_Plus = New Collection

If Aplicacion.ObtenerRsDAO(sql, rsP) Then
    Do While Not rsP.EOF
    
        Set cl_PP = New CLProdPrec
        
        cl_PP.codProd = rsP!limite_rango
        cl_PP.Precio = rsP!plus
        cl_PP.Tipo = rsP!tipo_pluss
        cl_PP.Objetivo = rsP!tipo_objetivo
        col_Plus.Add cl_PP
        
        rsP.MoveNext
    Loop
    Aplicacion.CerrarDAO rsP
End If
    
End Sub
Private Sub MeImpDatos()
End Sub
Private Sub MePrepararAgregar()

    Tollbar.Buttons(1).Value = tbrPressed
    Tollbar.Buttons(2).Value = tbrUnpressed
    
    Altas
    NuevaSeleccion
    
End Sub
Private Sub MePrepararAlterar()

    Tollbar.Buttons(2).Value = tbrPressed
    Tollbar.Buttons(1).Value = tbrUnpressed
    
    Modificacion
    NuevaSeleccion
    
End Sub
Public Sub Modificacion()

SetearBotonesAltas False
Modo = "MOD"
FrmAdmConcurso.caption = "Asignación de Productos -Modificacion y Bajas- "

End Sub
Private Sub NuevaSeleccion()
Dim i%

If Modo = "MOD" Then
    SetBotonesGeneral False
Else
    If chk.Value = 1 Then
        If MsgBox("Quiere salvar los cambios", vbOKCancel + vbQuestion, "ATENCION") = vbOK Then
            L_AltasDatos
        End If
    End If
End If
'Limpiar campos de pantallas
Set cl_ProvProd = New CLConc
Set col_Plus = New Collection

spr.MaxRows = 0
cboConcurso.ListIndex = -1
txtProv.Text = ""
mskPaga.Text = 0
msksec.Text = ""

    mskM.Text = 0
    mskP.Text = 0

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

FrmAdmConcurso.caption = Aplicacion.SeteoProceso(FrmAdmConcurso.caption)

Aplicacion.ComienzoTrans

MeLlenarObjeto


If cl_ProvProd.Update_Prod Then
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

FrmAdmConcurso.caption = Aplicacion.SeteoFin

Else
    MsgBox "Faltan Cargar Datos", vbExclamation + vbOKOnly, "ATENCION"
End If

End Sub
Private Sub MeCargarDatos()
Dim sql$
    
CondConsulta = ArmarCondicion

sqlGral$ = ""
sqlGral$ = sql$ & " SELECT C.id_concurso, C.cod_prov, C.secuencia,C.tipo,C.monto_prop,C.monto_paga " _
& " FROM estadis.concurso_prov C" _
& " " & CondConsulta
    
If Aplicacion.ObtenerRsDAO(sqlGral$, RS) Then
    txtCantReg.Text = Aplicacion.CantReg(RS)
    If txtCantReg.Text > 0 Then
        txtReg.Text = 1
        SetBotonesGeneral True
        MellenarPantalla
        MeSetearBotonesToolBar
    Else
        txtReg.Text = 0
    End If
End If

End Sub
Private Sub MeEliminar()

If MsgBox("Esta seguro de eliminar el registro", vbYesNo + vbExclamation, "ATENCION") = vbYes Then

'MeLlenarObjeto
cl_ProvProd.codConc = txtConc.Text
cl_ProvProd.CodProv = txtProv.Text
cl_ProvProd.sec = msksec.Text

FrmAdmConcurso.caption = Aplicacion.SeteoProceso(FrmAdmConcurso.caption)

Aplicacion.ComienzoTrans

If cl_ProvProd.Delete_COL_Prod Then
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

FrmAdmConcurso.caption = Aplicacion.SeteoFin
End If

End Sub
Private Sub MeLlenarObjeto()
Dim clProd As CLProdPrec
Dim i As Long, valor As Variant

cl_ProvProd.codConc = txtConc.Text
cl_ProvProd.CodProv = txtProv.Text
cl_ProvProd.sec = IIf(msksec.Text = "", -1, msksec.Text)

If optTipo(0).Value Then
    cl_ProvProd.Tipo = 0
    cl_ProvProd.MontoPaga = 0
    cl_ProvProd.MontoProp = 0
Else
    cl_ProvProd.Tipo = 1
    cl_ProvProd.MontoPaga = mskP.Text
    cl_ProvProd.MontoProp = mskM.Text
End If

Set cl_ProvProd.col_Plus = col_Plus
Set cl_ProvProd.col_producto = New Collection

For i = 1 To spr.MaxRows

    spr.GetText 1, i, valor
    Set clProd = New CLProdPrec
    If valor <> "" And Spread_FilaOcupada(spr, i) Then
        clProd.codProd = valor
        spr.GetText 3, i, valor
        clProd.Precio = valor
        AdicionarAColeccion clProd
        'If frmConcPlus.OptObjetivo(0).Value = True Then
             clProd.Objetivo = 0
        'Else
        '     clProd.Objetivo = 1
        'End If

    End If

Next

End Sub
Private Sub AdicionarAColeccion(cod As CLProdPrec)
Dim clProd As CLProdPrec
Dim Resp As Boolean

Resp = True

For Each clProd In cl_ProvProd.col_producto
    If cod.codProd = clProd.codProd Then
        Resp = False
        Exit For
    End If
Next

If Resp Then
    cl_ProvProd.col_producto.Add cod
End If

End Sub
Private Function L_TodoCargado() As Boolean
    L_TodoCargado = True
End Function
Private Sub MellenarPantalla()

Func.Func_SetearCboConLst cboConcurso, LstConcurso, RS!id_concurso

Func.Func_SetearCboConLst cboProv, lstProv, RS!cod_prov
msksec.Text = RS!secuencia

optTipo(RS!Tipo).Value = True
mskM.Text = RS!monto_prop
mskP.Text = RS!monto_paga

MeLlenarProductos

L_LLenarPlus

End Sub
Private Sub MeLlenarProductos()
Dim sql As String
Dim rsP As Recordset

sql = "SELECT P.cod_prod,P.descrip,CP.paga " _
& " FROM baires.producto P, estadis.concurso_d CP " _
& " WHERE  P.cod_prod = CP.cod_prod " _
& " and CP.id_concurso = '" & txtConc.Text & "' " _
& " and CP.cod_prov = '" & txtProv.Text & "' " _
& " and CP.secuencia = " & msksec.Text & " "

spr.MaxRows = 0
If Aplicacion.ObtenerRsDAO(sql, rsP) Then

    Do While Not rsP.EOF
        spr.MaxRows = spr.MaxRows + 1
        
        spr.SetText 1, spr.MaxRows, Trim(rsP!Cod_prod)
        spr.SetText 2, spr.MaxRows, Trim(rsP!Descrip)
        spr.SetText 3, spr.MaxRows, Trim(rsP!paga)
        rsP.MoveNext
    Loop
    Aplicacion.CerrarDAO rsP
End If

End Sub
Private Sub SetBotonesGeneral(valor As Boolean)
    
    Tollbar.Buttons(1).Enabled = Not valor
    Tollbar.Buttons(2).Enabled = Not valor
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(5).Enabled = Not valor
    
    Tollbar.Buttons(7).Enabled = valor
    Tollbar.Buttons(8).Enabled = valor
    Tollbar.Buttons(9).Enabled = valor
    Tollbar.Buttons(10).Enabled = valor

    Tollbar.Buttons(12).Enabled = valor
    Tollbar.Buttons(13).Enabled = valor

'habilitar frames
frCab.Enabled = Not valor
Spread.spread_LockGrilla spr, valor, 1, spr.MaxCols
botPlus.Enabled = valor
    If Not valor Then
        txtReg.Text = 0
        txtCantReg.Text = 0
    End If

End Sub
Private Function ArmarCondicion()
Dim Con$

Con$ = ""

If txtProv.Text <> "" Then
    Con$ = Con$ & " And c.cod_prov = '" & txtProv.Text & "' "
End If
If txtConc.Text <> "" Then
    Con$ = Con$ & " And c.id_concurso = '" & txtConc.Text & "' "
End If
If msksec.Text <> "" Then
    Con$ = Con$ & " And c.secuencia = " & msksec.Text
End If
If Con$ <> "" Then
    Con$ = " WHERE " & Mid(Con$, 5, Len(Con$))
End If

ArmarCondicion = Con$

End Function
Private Sub MePrepararMod()
    
    SeteoBotonesMod False

End Sub
Private Function MeReconsultar() As Integer
Dim sql$
Dim i%
    
If Aplicacion.ObtenerRsDAO(sqlGral$, RS) Then
        txtCantReg.Text = Aplicacion.CantReg(RS)
        If Val(txtReg.Text) > Val(txtCantReg.Text) Then
            txtReg.Text = txtCantReg.Text
        End If
        
        For i% = 1 To txtReg.Text - 1
            RS.MoveNext
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
    Tollbar.Buttons(7).Enabled = False
    Tollbar.Buttons(8).Enabled = False
    Tollbar.Buttons(9).Enabled = False
    Tollbar.Buttons(10).Enabled = False
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Enabled = True
ElseIf txtReg.Text = txtCantReg.Text Then
    Tollbar.Buttons(7).Enabled = True
    Tollbar.Buttons(8).Enabled = True
    Tollbar.Buttons(9).Enabled = False
    Tollbar.Buttons(10).Enabled = False
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Enabled = True
ElseIf txtReg.Text = 1 Then
    Tollbar.Buttons(7).Enabled = False
    Tollbar.Buttons(8).Enabled = False
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Enabled = True
Else
    Tollbar.Buttons(7).Enabled = True
    Tollbar.Buttons(8).Enabled = True
    Tollbar.Buttons(9).Enabled = True
    Tollbar.Buttons(10).Enabled = True
    Tollbar.Buttons(12).Enabled = True
    Tollbar.Buttons(13).Enabled = True
    
End If

End Sub
Private Sub SetearBotonesAltas(valor As Boolean)
'valor = true -> altas
'valor = false -> modif
    
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(15).Enabled = valor
    Tollbar.Buttons(16).Enabled = valor
    
    Tollbar.Buttons(17).Visible = Not valor 'False
    Tollbar.Buttons(18).Visible = Not valor 'False
    
    Tollbar.Buttons(5).Visible = Not valor 'False
    
    Tollbar.Buttons(7).Visible = Not valor 'False
    Tollbar.Buttons(8).Visible = Not valor 'False
    Tollbar.Buttons(9).Visible = Not valor 'False
    Tollbar.Buttons(10).Visible = Not valor 'False
    
    Tollbar.Buttons(12).Visible = Not valor 'False
    Tollbar.Buttons(13).Visible = Not valor 'False

    Tollbar.Buttons(18).Visible = Not valor 'False
    Tollbar.Buttons(19).Visible = Not valor 'False
    
    txtCantReg.Visible = Not valor 'False
    txtReg.Visible = Not valor 'False
    de.Visible = Not valor 'False

    frBot.Enabled = valor
    Frame2.Enabled = valor
    
    Spread.spread_LockGrilla spr, Not valor, 1, spr.MaxCols
    botPlus.Enabled = valor
    
End Sub
Private Sub SetearBotonesConsBases(valor As Boolean)
'valor = true -> altas
'valor = false -> modif
    
    Tollbar.Buttons(1).Visible = Not valor
    Tollbar.Buttons(2).Visible = Not valor
    
    Tollbar.Buttons(4).Visible = Not valor
    Tollbar.Buttons(15).Visible = Not valor
    Tollbar.Buttons(16).Visible = Not valor
    
    Tollbar.Buttons(17).Visible = Not valor 'False
    Tollbar.Buttons(18).Visible = Not valor 'False
    
    Tollbar.Buttons(5).Visible = Not valor 'False
    
    Tollbar.Buttons(7).Visible = valor  'False
    Tollbar.Buttons(8).Visible = valor  'False
    Tollbar.Buttons(9).Visible = valor  'False
    Tollbar.Buttons(10).Visible = valor  'False
    
    Tollbar.Buttons(12).Visible = Not valor 'False
    Tollbar.Buttons(13).Visible = Not valor 'False

    Tollbar.Buttons(18).Visible = Not valor 'False
    Tollbar.Buttons(19).Visible = Not valor 'False
    
    txtCantReg.Visible = Not valor 'False
    txtReg.Visible = Not valor 'False
    de.Visible = Not valor 'False

    frBot.Enabled = valor
    Frame2.Enabled = valor
    Spread.spread_LockGrilla spr, Not valor, 1, spr.MaxCols
    
End Sub
Public Sub PonerValores(cod As Variant, desc As String)
    spr.SetText 1, spr.MaxRows, Trim(cod)
    spr.SetText 2, spr.MaxRows, desc
    spr.SetText 3, spr.MaxRows, mskPaga.Text
'    If spr.ActiveRow = spr.MaxRows Then
       Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
'    End If
    If spr.MaxRows > 6 Then
        spr.Row = spr.MaxRows - 4
    Else
        spr.Row = spr.MaxRows
    End If
    spr.Col = 1
    spr.Position = SS_POSITION_UPPER_LEFT
    spr.Action = SS_ACTION_GOTO_CELL

DatoValido = True
End Sub
Private Sub SeteoBotonesMod(valor As Boolean)
     
    Tollbar.Buttons(4).Enabled = valor
    Tollbar.Buttons(5).Enabled = valor
    
    Tollbar.Buttons(7).Enabled = valor
    Tollbar.Buttons(8).Enabled = valor
    Tollbar.Buttons(9).Enabled = valor
    Tollbar.Buttons(10).Enabled = valor

    Tollbar.Buttons(12).Enabled = valor
    Tollbar.Buttons(13).Enabled = valor

    Tollbar.Buttons(15).Enabled = Not valor
    Tollbar.Buttons(16).Enabled = Not valor

    Tollbar.Buttons(18).Enabled = valor
    Tollbar.Buttons(19).Enabled = valor
'habilitar o des frames y/o campos
Spread.spread_LockGrilla spr, valor, 1, spr.MaxCols
frBot.Enabled = Not valor
Frame2.Enabled = Not valor
'botPlus.Enabled = Not valor
End Sub
Private Sub botHelpProd_Click()
Dim cod As String
Dim desc As String, CodProv As String
Dim sql As String

 If txtProv.Text <> "" Then
    If spr.MaxRows > 0 Then
    If txtProv.Text <> "N" Then
        sql = "Select cod_prod,descrip from baires.producto " _
        & " Where cod_prov = '" & txtProv.Text & "' "
        CodProv = txtProv.Text
    Else
        sql = "Select cod_prod,descrip from baires.producto " _
        & " Where motivo_discon <> 'N' "
        CodProv = ""
    End If
    If frmHelpProd.MuestraHlp(cod, desc, "Producto", sql, CodProv) = vbOK Then
       spr.SetText 1, spr.ActiveRow, cod
       spr.SetText 2, spr.ActiveRow, desc
       If spr.ActiveRow = spr.MaxRows Then
          Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
       End If
       DatoValido = True
    End If
    End If
Else
    MsgBox "Elija primero un Proveedor", vbOKOnly + vbExclamation, "ATENCION"
End If
spr.SetFocus

End Sub
Private Sub botPlus_Click()

    frmConcPlus.ProductosMonitoreo col_Plus
    
End Sub
Private Sub cboConcurso_Click()
Dim sql As String

txtConc.Text = LstConcurso.List(cboConcurso.ListIndex)

sql = " SELECT P.cod_prov,P.descrip FROM estadis.concurso_PROV C,baires.proveedor P " _
& " WHERE P.cod_prov=C.cod_prov And C.id_concurso = '" & txtConc.Text & "' " _
& " UNION " _
& " SELECT CP.cod_prov, 'DISCONTINUADOS' AS Descrip " _
& " FROM estadis.concurso_prov CP " _
& " WHERE  CP.cod_prov ='N' " _
& " and CP.id_concurso = '" & txtConc.Text & "' "


FuncCbos_LlenarCboLst cboProv, lstProv, sql

End Sub
Private Sub cboConcurso_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    cboConcurso.ListIndex = -1
    txtProv.Text = ""
ElseIf KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub cboProv_Click()
txtProv.Text = lstProv.List(cboProv.ListIndex)
End Sub
Private Sub cboProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    cboProv.ListIndex = -1
ElseIf KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 4 Then
    Select Case KeyCode
        Case 78 'Nueva sel
            Call TollBar_ButtonClick(Tollbar.Buttons(4))
        Case 71 'Guardar
            If Tollbar.Buttons(12).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(15))
            End If
        Case 66 'Buscar
            If Modo = "MOD" Then
            Call TollBar_ButtonClick(Tollbar.Buttons(5))
            End If
        Case 83 'Salir
            Call TollBar_ButtonClick(Tollbar.Buttons(19))
    End Select
    If Modo = "MOD" And Val(txtCantReg.Text) > 0 Then
    Select Case KeyCode
        Case 37 'Izq
            If Tollbar.Buttons(5).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(8))
            End If
        Case 38 'Arriba
            Call TollBar_ButtonClick(Tollbar.Buttons(10))
        Case 40 'Abajo
            Call TollBar_ButtonClick(Tollbar.Buttons(7))
        Case 39 'Der
            If Tollbar.Buttons(6).Enabled Then
            Call TollBar_ButtonClick(Tollbar.Buttons(9))
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
Dim sql As String

Top = 700
Left = 800
Width = 7400
Height = 6500

FrmAdmConcurso.caption = "Asignación de Productos -Modificacion y Bajas- "

sql = " SELECT id_concurso,descrip FROM estadis.concurso_H Where anio_alta > 2001 "
sql = sql & " ORDER BY anio_alta desc, mes desc , id_concurso "

FuncCbos_LlenarCboLst cboConcurso, LstConcurso, sql

Modo = "MOD"
NuevaSeleccion
DatoValido = True

End Sub

Private Sub mskM_LostFocus()

If Not IsNumeric(mskM.Text) Then
    mskM.Text = 0
End If

End Sub


Private Sub mskP_LostFocus()
If Not IsNumeric(mskP.Text) Then
    mskP.Text = 0
End If

End Sub


Private Sub mskPaga_LostFocus()

If Not IsNumeric(mskPaga.Text) Then
    mskPaga.Text = 0
End If

End Sub

Private Sub spr_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If Mode = 1 Then
    ModoEdit = Mode
End If
End Sub
Private Sub spr_KeyPress(KeyAscii As Integer)
Dim sql As String
Dim desc As String
Dim cod As Variant
Dim st As Variant

If KeyAscii = 13 And ModoEdit = 1 Then
    
    desc = ""
    ModoEdit = 0
    spr.GetText 1, spr.ActiveRow, cod
    
    sql = "SELECT descrip "
    sql = sql & " FROM  baires.producto "
    sql = sql & " where cod_prod = " & cod
    sql = sql & " And cod_prov = '" & txtProv.Text & "'"
    Select Case spr.ActiveCol
        Case 1
            DatoValido = Func_ObtenerDesc(sql, desc)
            spr.SetText 2, spr.ActiveRow, desc
            If DatoValido Then
                spr.Col = 3
                spr.Row = spr.ActiveRow
                spr.Action = 0
                spr.Action = 1
            End If

        Case 2
        Case 3
            spr.GetText 3, spr.ActiveRow, st
            DatoValido = Func_ObtenerDesc(sql, desc)
            If DatoValido And spr.ActiveRow = spr.MaxRows Then
                Call Toolbar1_ButtonClick(Toolbar1.Buttons(1))
                spr.Col = 1
                spr.Row = spr.MaxRows
                spr.Action = 0
                spr.Action = 1
            End If

    End Select
            spr.TopRow = spr.TopRow - 6
End If

End Sub
Private Sub spr_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim sql As String
Dim desc As String
Dim cod As Variant
Dim st As Variant

If ModoEdit = 1 Then
    
    desc = ""
    ModoEdit = 0
    spr.GetText 1, Row, cod
    
    sql = "SELECT descrip "
    sql = sql & " FROM  baires.producto "
    sql = sql & " where cod_prod = " & cod
    sql = sql & " And cod_prov = '" & txtProv.Text & "'"
    
    Select Case spr.ActiveCol
        Case 1
            DatoValido = Func_ObtenerDesc(sql, desc)
            spr.SetText 2, spr.ActiveRow, desc
        Case 2
        Case 3
            spr.GetText 3, spr.ActiveRow, st
            DatoValido = Func_ObtenerDesc(sql, desc)
    End Select
    
End If

Cancel = Not DatoValido

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
         Func_MoverPrimero RS, pos
    Case "b"
         saltear = False
        Func_MoverAnterior RS, pos
    Case "c"
         saltear = False
        Func_MoverSiguiente RS, pos
    Case "e"
         saltear = False
        Func_MoverUltimo RS, pos
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
    Case "o"
        MePrepararAgregar
    Case "p"
        MePrepararAlterar
End Select

If Not saltear Then
    txtReg.Text = pos
    MellenarPantalla
    MeSetearBotonesToolBar
End If

End Sub
Private Sub L_DatosGrilla()
Dim i
Dim nro As Integer

On Error GoTo DG:

If nro > 0 Then
    RS.MoveFirst
    For i = 1 To nro - 1
        RS.MoveNext
    Next
    MellenarPantalla
    txtReg.Text = nro
    MeSetearBotonesToolBar

DG:
    Exit Sub
End If

End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim valor As Variant

Select Case Button.Key
    Case "A"
        If Spread_FilaOcupada(spr, spr.MaxRows) Then
           Spread_AddRow spr
        End If
    Case "B"
        Spread_DelOneRow spr, spr.ActiveRow
        DatoValido = True
    Case "C"
        spr.MaxRows = 0
    Case "D"
        
End Select
End Sub
