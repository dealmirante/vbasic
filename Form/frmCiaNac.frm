VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmCiaNac 
   Caption         =   "Información por Nacionalidad"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3855
   ScaleWidth      =   7590
   Begin VB.Frame frEzeA0 
      Height          =   2985
      Left            =   60
      TabIndex        =   7
      Top             =   1380
      Width           =   7440
      Begin FPSpread.vaSpread spr 
         Height          =   2700
         Left            =   120
         OleObjectBlob   =   "frmCiaNac.frx":0000
         TabIndex        =   8
         Top             =   180
         Width           =   7155
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   6030
      Begin VB.Label labAerop 
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
         Height          =   270
         Left            =   1470
         TabIndex        =   6
         Top             =   1005
         Width           =   2700
      End
      Begin VB.Label labCia 
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
         Height          =   270
         Left            =   1470
         TabIndex        =   5
         Top             =   630
         Width           =   2700
      End
      Begin VB.Label labFechas 
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
         Height          =   270
         Left            =   1485
         TabIndex        =   4
         Top             =   285
         Width           =   2700
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aeropuerto"
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
         Left            =   135
         TabIndex        =   3
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cia Aérea"
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
         Left            =   135
         TabIndex        =   2
         Top             =   615
         Width           =   1185
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
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmCiaNac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Sdep As String
Dim fchD As String
Dim fchH As String
Dim cia As String
Dim codCia As String

Public Sub MostrarInfo(FD As String, FH As String, CodA As Integer, CA As String, AE As String)

fchD = FD
fchH = FH
cia = CA
codCia = CodA
Sdep = AE

Me.Show


End Sub

Private Sub Form_Load()

Width = Screen.Width * 0.8
Height = Screen.Height * 0.7

labFechas.caption = fchD & IIf(fchH = "", "", " - " & fchH)

Select Case Sdep
    Case "AEP"
        labAerop.caption = "AEROPARQUE "
    Case "INTA"
        labAerop.caption = "INTERNACIONAL 'A' "
    Case "INTB"
        labAerop.caption = "INTERNACIONAL 'B' "
End Select

labCia.caption = cia

L_Refrescar

End Sub

Private Sub L_Refrescar()
Dim sql As String
Dim sqlX As String
Dim rs As Recordset

On Error GoTo ErrNac:

Call Aplicacion.SeteoProceso("")

If codCia <> -1 Then
    sql = " SELECT "
    sql = sql & " P.descrip, "
    sql = sql & " sum(cant_tickets) cant_t, "
    sql = sql & " sum(importe) imp, "
    sql = sql & " sum(cant_pax) cant_p "
    sql = sql & "FROM " & funcLocal_Vista("pax_espigon", Year(fchD))
    sql = sql & " V, ventas.paises P "
    sql = sql & " WHERE nacionalidad=cod_pais "
    sql = sql & " And fch_vta BETWEEN " & func_ToDate(fchD) & " And " & IIf(fchH = "", func_ToDate(fchD), func_ToDate(fchH))
    sql = sql & " And cod_cia_aerea = " & codCia & ""
    sql = sql & " And cod_sdep  = '" & Sdep & "'"
    sql = sql & "group by p.descrip "
    sql = sql & " order by p.descrip "
Else
    sql = " SELECT "
    sql = sql & " P.descrip, "
    sql = sql & " sum(cant_tickets) cant_t, "
    sql = sql & " sum(importe) imp, "
    sql = sql & " sum(cant_pax) cant_p "
    sql = sql & "FROM " & funcLocal_Vista("pax_espigon", Year(fchD))
    sql = sql & " V, ventas.paises P "
    sql = sql & " WHERE nacionalidad=cod_pais "
    sql = sql & " And fch_vta BETWEEN " & func_ToDate(fchD) & " And " & IIf(fchH = "", func_ToDate(fchD), func_ToDate(fchH))
    sql = sql & " And cod_sdep  = '" & Sdep & "'"
    sql = sql & "group by p.descrip "
    sql = sql & " order by p.descrip "
End If


If Aplicacion.ObtenerRsDAO(sql, rs) Then
    spr.MaxRows = 0
    Do While Not rs.EOF
        spr.MaxRows = spr.MaxRows + 1
        
        spr.SetText 1, spr.MaxRows, Trim(rs!Descrip)
        spr.SetText 2, spr.MaxRows, str(rs!imp)
        spr.SetText 3, spr.MaxRows, str(rs!cant_t)
        spr.SetText 4, spr.MaxRows, str(rs!cant_p)
        
        rs.MoveNext
    Loop
    Aplicacion.CerrarDAO rs
End If

Spread_TotalesGrillas spr, spr.MaxCols - 3, 2

ErrNac:
    Call Aplicacion.SeteoFin

End Sub

