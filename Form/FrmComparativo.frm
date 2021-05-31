VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSACAL70.OCX"
Begin VB.Form FrmComparativo 
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3600
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3990
      Begin VB.CommandButton CmdProcesar 
         Caption         =   "Procesar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   810
         TabIndex        =   3
         Top             =   3135
         Width           =   1440
      End
      Begin VB.CommandButton CmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   2
         Top             =   3120
         Width           =   1440
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   2265
         Left            =   75
         TabIndex        =   1
         Top             =   780
         Width           =   3855
         _Version        =   458752
         _ExtentX        =   6800
         _ExtentY        =   3995
         _StockProps     =   1
         Year            =   1998
         Month           =   9
         Day             =   1
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridCellEffect  =   2
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridFontColor   =   16711680
         GridLinesColor  =   -2147483632
         MonthLength     =   0
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleFontColor  =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FECHA DEL PROCESO"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   165
         TabIndex        =   4
         Top             =   270
         Width           =   3675
      End
   End
End
Attribute VB_Name = "FrmComparativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim fecha As String
Public Sub MeComparar()

Dim SQL As String

SQL = ""
SQL = SQL & " BEGIN "
SQL = SQL & " estadis.Procesos_Estadisticos.crea_pax_at_vs_vol("
SQL = SQL & func_ToDate(fecha) & ");"
SQL = SQL & " END; "


Aplicacion.ComienzoTrans


If Aplicacion.EjecutarDAO(SQL) Then
   Aplicacion.TerminarConExitoTrans
 Else
   Aplicacion.TerminarConErrorTrans
End If
    Unload Me
End Sub


Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdProcesar_Click()

   fecha = Format$(Calendar1.Value, FTOFECHA)
   MeComparar

   Call Aplicacion.EjecutarDAO("begin estadis.procesos_estadisticos_mes.INDICADORES_pax_webDB; end;")
   
End Sub


Private Sub Form_Load()
Calendar1.Day = Day(Now)
Calendar1.Year = Year(Now)
Calendar1.Month = Month(Now)
End Sub

