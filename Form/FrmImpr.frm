VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.0#0"; "CRYSTL32.OCX"
Begin VB.Form FrmImprFrom 
   Caption         =   "Impresion de Formularios"
   ClientHeight    =   2820
   ClientLeft      =   2610
   ClientTop       =   1575
   ClientWidth     =   3675
   ControlBox      =   0   'False
   Icon            =   "FrmImpr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2820
   ScaleWidth      =   3675
   Begin VB.Frame Frame1 
      Height          =   2715
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   3555
      Begin MSMask.MaskEdBox makPD 
         Height          =   315
         Left            =   1050
         TabIndex        =   9
         Top             =   1710
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         _Version        =   327680
         PromptChar      =   "_"
      End
      Begin VB.OptionButton optImp 
         Caption         =   "Por Impresora"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   690
         Width           =   1560
      End
      Begin VB.OptionButton optImp 
         Caption         =   "Por Pantalla"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1560
      End
      Begin Crystal.CrystalReport rptForm 
         Left            =   3000
         Top             =   2355
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   327682
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton botEjecutar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   315
         Index           =   0
         Left            =   2070
         TabIndex        =   3
         Top             =   1785
         Width           =   1260
      End
      Begin VB.CommandButton botEjecutar 
         Caption         =   "&Aceptar"
         Default         =   -1  'True
         Height          =   315
         Index           =   1
         Left            =   2055
         TabIndex        =   2
         Top             =   1200
         Width           =   1260
      End
      Begin VB.TextBox intCopias 
         Height          =   285
         Left            =   1035
         TabIndex        =   1
         Text            =   "1"
         Top             =   1215
         Width           =   705
      End
      Begin MSMask.MaskEdBox makPH 
         Height          =   315
         Left            =   1050
         TabIndex        =   10
         Top             =   2100
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         _Version        =   327680
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Pag Hasta"
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
         Left            =   75
         TabIndex        =   8
         Top             =   2145
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Pag Desde"
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
         Left            =   90
         TabIndex        =   7
         Top             =   1785
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad de copias"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   1170
         Width           =   870
      End
   End
End
Attribute VB_Name = "FrmImprFrom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lFecha As String
Dim lFechaH As String

Dim lCia As Integer
Dim lSdep As String
Dim lTipo As String
Dim lMonto As Double
Dim lLegajo As Long

Dim Info As String

Private Sub MeImprimirCia()
Dim i%

On Error GoTo ErrPrint:


FrmImprFrom.caption = Aplicacion.SeteoProceso(FrmImprFrom.caption)

rptForm.ReportFileName = RutaReportes
rptForm.ReportFileName = RutaReportes & "atvsvol.rpt"
rptForm.Formulas(0) = "flaSDep = '" & lSdep & "'"
rptForm.Formulas(1) = "flaCia = " & lCia
rptForm.Formulas(2) = "flaFecha = DATE(" & Year(lFecha) & "," & Month(lFecha) & "," & Day(lFecha) & ")"
rptForm.Formulas(3) = "flatipo = '" & lTipo & "'"

rptForm.Action = True
DoEvents

FrmImprFrom.caption = Aplicacion.SeteoFin

Exit Sub
ErrPrint:
     FrmImprFrom.caption = Aplicacion.SeteoFin
     Exit Sub
End Sub

Private Sub MeImprimirMod()
Dim i%

'On Error GoTo ErrPrint:


FrmImprFrom.caption = Aplicacion.SeteoProceso(FrmImprFrom.caption)

rptForm.ReportFileName = RutaReportes
rptForm.ReportFileName = RutaReportes & "modelos.rpt"
rptForm.Formulas(0) = "flaValor = " & lMonto & ""


rptForm.Action = True
DoEvents

If lTipo = "I" Then
    rptForm.ReportFileName = RutaReportes & "modrubro.rpt"
    rptForm.Formulas(0) = "flaValor = " & lMonto & ""
    
    rptForm.Action = True
    DoEvents

    rptForm.ReportFileName = RutaReportes & "moddiaru.rpt"
    rptForm.Formulas(0) = "flaValor = " & lMonto & ""
    
    rptForm.Action = True
    DoEvents

End If

FrmImprFrom.caption = Aplicacion.SeteoFin

Exit Sub
ErrPrint:
     FrmImprFrom.caption = Aplicacion.SeteoFin
     Exit Sub
End Sub


Private Sub MeImprimirTot()
Dim i%

On Error GoTo ErrPrint:


FrmImprFrom.caption = Aplicacion.SeteoProceso(FrmImprFrom.caption)

rptForm.ReportFileName = RutaReportes
rptForm.ReportFileName = RutaReportes & "totespigon.rpt"
rptForm.Formulas(0) = "flaFechaD = DATE(" & Year(lFecha) & "," & Month(lFecha) & "," & Day(lFecha) & ")"
rptForm.Formulas(1) = "flaFechaH = DATE(" & Year(lFechaH) & "," & Month(lFechaH) & "," & Day(lFechaH) & ")"

rptForm.Action = True
DoEvents

FrmImprFrom.caption = Aplicacion.SeteoFin

Exit Sub
ErrPrint:
     FrmImprFrom.caption = Aplicacion.SeteoFin
     Exit Sub
End Sub

Private Sub MeImprimirAus()
Dim i%

On Error GoTo ErrPrint:


FrmImprFrom.caption = Aplicacion.SeteoProceso(FrmImprFrom.caption)

rptForm.ReportFileName = RutaReportes
rptForm.ReportFileName = RutaReportes & "ausencia.rpt"
If lFecha <> "" Then
    rptForm.ParameterFields(0) = "fechaD;" & " DATE(" & Year(CDate(lFecha)) & "," & Month(CDate(lFecha)) & "," & Day(CDate(lFecha)) & ")" & ";true)"
Else
    rptForm.ParameterFields(0) = "fechaD;date(1900,01,01);true)"
End If
If lFechaH <> "" Then
    rptForm.ParameterFields(1) = "fechaH;" & " DATE(" & Year(CDate(lFechaH)) & "," & Month(CDate(lFechaH)) & "," & Day(CDate(lFechaH)) & ")" & ";true)"
Else
    rptForm.ParameterFields(1) = "fechah;date(1900,01,01);true)"
End If
rptForm.ParameterFields(2) = "legajo;" & lLegajo & ";true)"

rptForm.Action = True
DoEvents

FrmImprFrom.caption = Aplicacion.SeteoFin

Exit Sub
ErrPrint:
     FrmImprFrom.caption = Aplicacion.SeteoFin
     Exit Sub
End Sub



Private Sub MeImprimirTotCia()
Dim i%

On Error GoTo ErrPrint:


FrmImprFrom.caption = Aplicacion.SeteoProceso(FrmImprFrom.caption)

rptForm.ReportFileName = RutaReportes
rptForm.ReportFileName = RutaReportes & "totcia.rpt"
rptForm.Formulas(0) = "flaFechaD = DATE(" & Year(lFecha) & "," & Month(lFecha) & "," & Day(lFecha) & ")"
rptForm.Formulas(1) = "legajo = DATE(" & Year(lFechaH) & "," & Month(lFechaH) & "," & Day(lFechaH) & ")"

rptForm.Action = True
DoEvents

FrmImprFrom.caption = Aplicacion.SeteoFin

Exit Sub
ErrPrint:
     FrmImprFrom.caption = Aplicacion.SeteoFin
     Exit Sub
End Sub



Public Sub TratarImpresionCia(Fecha As String, Sdep As String, cia As Integer, Tipo As String)

lFecha = Fecha
lSdep = Sdep
lCia = cia
lTipo = Tipo

Info = "CIA"

Me.Show 1


End Sub

Public Sub TratarImpresionModelos(monto As Double, Tipo As String)

lMonto = monto
lTipo = Tipo
Info = "MOD"

Me.Show 1

End Sub


Public Sub TratarImpresionTot(fechaD As String, fechaH As String)

lFecha = fechaD
lFechaH = fechaH

Info = "TOT"

Me.Show 1


End Sub

Public Sub TratarImpresionAusent(fechaD As String, fechaH As String, leg As String)

lFecha = IIf(fechaD = "  -  -    ", "", fechaD)
lFechaH = IIf(fechaH = "  -  -    ", "", fechaH)

lLegajo = IIf(leg = "", 0, leg)

Info = "AUS"

Me.Show 1


End Sub


Public Sub TratarImpresionTotCia(fechaD As String, fechaH As String)

lFecha = fechaD
lFechaH = fechaH

Info = "TOTCia"

Me.Show 1


End Sub

Private Sub botEjecutar_Click(Index As Integer)

If Index = 0 Then
    Unload Me
Else
    Select Case Info
     Case "AUS"
        MeImprimirAus
     Case "CIA"
        MeImprimirCia
     Case "TOT"
        MeImprimirTot
     Case "MOD"
        MeImprimirMod
     Case "TOTCia"
        MeImprimirTotCia
    
    End Select
    
'    Unload Me
End If

End Sub
Private Sub Form_Load()
Dim User$

Top = 2000
Left = 3000

rptForm.Connect = "ODBC;UID=" & Aplicacion.username & ";PWD=" & Aplicacion.password & ";DATABASE=INTER;"

rptForm.PrinterCopies = 1
rptForm.Destination = 0

makPD.Text = 0
makPH.Text = 0

rptForm.PrinterStopPage = 0
rptForm.PrinterStartPage = 0

End Sub


Private Sub intCopias_LostFocus()

If IsNumeric(intCopias.Text) Then
    rptForm.PrinterCopies = intCopias.Text
Else
    intCopias.Text = 1
    rptForm.PrinterCopies = intCopias.Text
End If

End Sub




Private Sub makPD_LostFocus()

rptForm.PrinterStartPage = makPD.Text
If makPD.Text = 0 Then
    makPH.Text = 0
    rptForm.PrinterStopPage = 0
End If

End Sub


Private Sub makPH_LostFocus()

rptForm.PrinterStopPage = makPH.Text
If makPH.Text = 0 Then
    makPD.Text = 0
    rptForm.PrinterStartPage = 0
End If


End Sub


Private Sub optImp_Click(Index As Integer)

rptForm.Destination = Index ' 0 pantalla
                            '1 impresora
End Sub


