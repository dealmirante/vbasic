VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmGriPersonal 
   Caption         =   "Form1"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin FPSpread.vaSpread sprPer 
      Height          =   3090
      Left            =   60
      OleObjectBlob   =   "frmGriPersonal.frx":0000
      TabIndex        =   0
      Top             =   405
      Width           =   5625
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TXT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5295
      TabIndex        =   2
      Top             =   30
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Height          =   360
      Left            =   75
      TabIndex        =   1
      Top             =   30
      Width           =   5220
   End
End
Attribute VB_Name = "frmGriPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Esp As String
Dim Porc As Single
Dim spr As Control

Private Function L_CalcularPorc(rub As String) As String
Dim Pc As Single
Dim i, valor As Variant

    Pc = 0
    For i = 1 To spr.MaxRows
        spr.GetText 2, i, valor
        If valor = rub Then
            spr.GetText 4, i, valor
            'Pc = (Porc * 40 / 100) + (Porc * 60 * Val(valor) / 10000)
            Pc = (Porc * 40 / 100) + Val(valor)
            Exit For
        End If
    Next
    If Pc = 0 Then
        Pc = (Porc * 40 / 100)
    End If
    L_CalcularPorc = str(Pc)
End Function

Private Sub L_TratarPersonal()
Dim rs As Recordset, rsCod As Recordset, fila As Integer
Dim i, valor As Variant, cod As String, desc As String
Dim rub As String, sql As String, sqla As String
Dim Pc As Single

    sql = "SELECT cod_equipo,grupo,rubros  FROM estadis.equipos " _
        & "WHERE " _
        & " cod_sdep  = '" & Esp & "'"
    Call Aplicacion.ObtenerRsDAO(sql, rsCod)
    fila = sprPer.MaxRows + 1
    Do While Not rsCod.EOF
        sql = "SELECT legajo FROM estadis.Persona_Equipos " _
        & " WHERE cod_equipo = " & rsCod!cod_equipo
            
        If Aplicacion.ObtenerRsDAO(sql, rs) Then
            If Aplicacion.CantReg(rs) > 0 Then
                Do While Not rs.EOF
                    sprPer.MaxRows = sprPer.MaxRows + 1
                    sprPer.SetText 1, fila, str(rs!Legajo)
                    
                    sqla = "SELECT E.legajo,E.legajo,Apellido || ', ' || nombre  as descrip " _
                    & "FROM personal.empleado E " _
                    & " WHERE E.legajo = " & rs!Legajo
                    
                    Call Func_ObtenerDesc(sqla, desc)
                    sprPer.SetText 2, fila, Trim(desc)
                    
                    sprPer.SetText 3, fila, L_CalcularPorc(rsCod!Rubros)
                    
                    fila = fila + 1
                    rs.MoveNext
                Loop
            Else
            End If
            Aplicacion.CerrarDAO rs
            End If

        rsCod.MoveNext
    Loop

End Sub



Public Sub MostrarPersonal(sEsp As String, sPorc As Single, cspr As Control)

Esp = sEsp
Porc = sPorc
Set spr = cspr
sprPer.MaxRows = 0
Select Case Esp
    Case "INTA"
        Label1.caption = "Internacional A"
    Case "INTB"
        Label1.caption = "Internacional B"
    Case "AEP"
        Label1.caption = "Aeroparque"
End Select
L_TratarPersonal

Me.Show 1


End Sub


Public Sub MostrarPersonalCIA(sEspA As String, sPorcA As Single, csprA As Control, _
                              sEspB As String, sPorcB As Single, csprB As Control, _
                              sEspE As String, sPorcE As Single, csprE As Control)

sprPer.MaxRows = 0
Label1.caption = "COMPAÑIA"
Esp = sEspA
Porc = sPorcA
Set spr = csprA
If spr.MaxRows > 0 Then
    L_TratarPersonal
End If
Esp = sEspB
Porc = sPorcB
Set spr = csprB
If spr.MaxRows > 0 Then
    L_TratarPersonal
End If
Esp = sEspE
Porc = sPorcE
Set spr = csprE
If spr.MaxRows > 0 Then
    L_TratarPersonal
End If

Me.Show 1

End Sub


Private Sub Command1_Click()
Dim NOMBRE As String
Dim linea As String, i
Dim valor As Variant

NOMBRE = frmDir.NombreArch("txt")
If NOMBRE <> "" Then
    Open NOMBRE For Output Access Write Lock Write As #1
    
        For i = 1 To sprPer.MaxRows
            linea = ""
            sprPer.GetText 1, i, valor
            linea = Format$(valor, "00000")
            'linea = linea & " "
            sprPer.GetText 3, i, valor
            linea = linea & L_SinPunto(valor)
            
            Print #1, linea
            
        Next
        
    Close #1
End If

End Sub

Private Function L_SinPunto(dato As Variant) As String
Dim pos
Dim Ent As String, Deci As String


pos = InStr(1, dato, ".")
If pos = 0 Then
    L_SinPunto = Format$(dato, "000") & "00"
Else
    Ent = Format$(Left(dato, pos - 1), "000")
    Deci = Format$(Mid(dato, pos + 1), "#0")
    L_SinPunto = Ent & Deci
End If

End Function


Private Sub Form_Load()

Width = Screen.Width * 0.61
Height = Screen.Height * 0.57
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height - 500) / 2


End Sub


