Attribute VB_Name = "mPrincipal"
Option Explicit

Public Sub Main()
Const iMaxIntentos = 3
Dim iIntentos As Integer
Dim bSalir As Boolean
Dim bConecto As Boolean, bConectoP As Boolean
Dim Ruta As String
Dim Clave As String
Dim Seccion As String
Dim cant As Integer
Dim StrRet As String
Dim defol As String

Ruta = "win.ini"
Seccion = "Intl"
Clave = "sDecimal"
cant = 2
StrRet = String(1, " ")

Call SaveSetting("Aplicaciones", "Estadis", "Anio", 364)

If GetSetting("Aplicaciones", "Estadis", "CtrolEST", "NoEsta") = "NoEsta" Then
    Call SaveSetting("Aplicaciones", "Estadis", "CtrolEST", "SI")
    Call SaveSetting("Aplicaciones", "Estadis", "RutaFotos", App.Path & "\Fotos\")
    Call SaveSetting("Aplicaciones", "Estadis", "RutaReports", App.Path & "\Reportes\")
    Call SaveSetting("Aplicaciones", "Estadis", "Usuario", "")
    Call SaveSetting("Aplicaciones", "Estadis", "Anio", 364)
    RutaFotos = App.Path & "\Fotos\"
    RutaReportes = App.Path & "\Reportes\"
    CantAnio = 364
Else
    RutaFotos = GetSetting("Aplicaciones", "Estadis", "RutaFotos", App.Path & "\Fotos\")
    RutaReportes = GetSetting("Aplicaciones", "Estadis", "RutaReports", App.Path & "\Reportes\")
    CantAnio = GetSetting("Aplicaciones", "Estadis", "Anio", 364)
End If

Call GetPrivateProfileString(Seccion, Clave, defol, StrRet, cant, Ruta)

If StrRet = "." Then

'Crear objeto aplicacion
Set Aplicacion = New CAplicacion

'Mostrar pantalla de login
iIntentos = 1
bSalir = False
bConecto = False
frmMDI.Show

frmPrincipal.Show
DoEvents

Aplicacion.CaptionPross = "Procesando.."
While iIntentos <= iMaxIntentos And Not bConecto

    If FrmLogon.PedirUsuario Then 'pedir usuario muestra la pantalla de log y si el usuario continúa'

       If Aplicacion.conectarDAO Then
          bConecto = True
       '  Call Aplicacion.SeteoProceso("")
          Unload FrmLogon
          iIntentos = 1
          Call MePerfilUser

          DoEvents
          Call SaveSetting("Aplicaciones", "Estadis", "Usuario", Aplicacion.username)
          Aplicacion.anio = CantAnio
          frmPrincipal.Hide

          MeLlenarDSLocal

          DoEvents

          L_setearopciones
          frmTM.TM.Enabled = True

       Else
          iIntentos = iIntentos + 1
       End If

    Else
      Unload FrmLogon
      Unload frmMDI
      Exit Sub
    End If

Wend

If iIntentos > iMaxIntentos Then
  Unload FrmLogon
  Unload frmMDI
  Exit Sub
End If
Else
    MsgBox "Ud. no puede entrar al Sistema . Antes debe configurar en el PANEL DE CONTROL, en la opción INTERNACIONAL. El Punto (.) como separador decimal y la Coma (,) como sepearador de miles. Si Ud. no puede hacerlo, llamar a Sistemas.", vbExclamation + vbOKOnly, "ATENCION"

End If
    

End Sub
Private Sub L_setearopciones()
If Aplicacion.Perfil = "TODOS" Then
        frmMDI.MnuGral(0).Visible = True
        frmMDI.MnuGral(1).Visible = True
        frmMDI.MnuGral(2).Visible = True
        frmMDI.MnuGral(5).Visible = True
        frmMDI.MnuGral(6).Visible = True
    
    Select Case Aplicacion.Nivel
    Case 0
        frmMDI.MnuGral(7).Visible = True
        
    End Select
Else
    'PARA TODOS
    frmMDI.mnuEst(16).Visible = False
    frmMDI.mnuvs(13).Visible = False
            
    frmMDI.MnuGral(0).Visible = True
    frmMDI.MnuGral(1).Visible = True
    frmMDI.MnuGral(2).Visible = True
    frmMDI.MnuGral(5).Visible = True
    frmMDI.MnuGral(6).Visible = True

Select Case Aplicacion.Nivel
    Case 0
        frmMDI.MnuGral(7).Visible = True
            frmMDI.mnuAdm(8).Visible = False
    Case 2
        frmMDI.MnuGral(7).Visible = True
            frmMDI.mnuAdm(0).Visible = False
            frmMDI.mnuAdm(1).Visible = False
            frmMDI.mnuAdm(2).Visible = False
            frmMDI.mnuAdm(3).Visible = False
            frmMDI.mnuAdm(4).Visible = False
            frmMDI.mnuAdm(8).Visible = False
        frmMDI.mnuCom(4).Visible = False
        frmMDI.mnuCap(0).Visible = False
        frmMDI.mnuCap(1).Visible = False
    Case 10
        frmMDI.mnuCom(4).Visible = False
End Select
End If
End Sub


Private Sub MeLlenarDSLocal()
DSLoc(1).Dep = "AEP"
DSLoc(1).Sdep = "AEP"
DSLoc(1).locales(1) = "L090"
DSLoc(1).SLocales(1) = 0
DSLoc(1).locales(2) = "L140"
DSLoc(1).SLocales(2) = 0
DSLoc(1).locales(3) = "L141"
DSLoc(1).SLocales(3) = 1


DSLoc(2).Dep = "EZE"
DSLoc(2).Sdep = "INTA"
DSLoc(2).locales(1) = "L050"
DSLoc(2).SLocales(1) = 0
DSLoc(2).locales(2) = "L051"
DSLoc(2).SLocales(2) = 1
DSLoc(2).locales(3) = "L060"
DSLoc(2).SLocales(3) = 0
DSLoc(2).locales(4) = "L061"
DSLoc(2).SLocales(4) = 1

DSLoc(2).locales(5) = "L220"
DSLoc(2).SLocales(5) = 0

DSLoc(2).locales(6) = "L221"
DSLoc(2).SLocales(6) = 1
DSLoc(2).locales(7) = "L222"
DSLoc(2).SLocales(7) = 2
DSLoc(2).locales(8) = "L223"
DSLoc(2).SLocales(8) = 3
DSLoc(2).locales(9) = "L224"
DSLoc(2).SLocales(9) = 4


DSLoc(9).Dep = "EZE"
DSLoc(9).Sdep = "INTAL"

DSLoc(10).Dep = "EZE"
DSLoc(10).Sdep = "INTAS"

DSLoc(3).Dep = "EZE"
DSLoc(3).Sdep = "INTB"
DSLoc(3).locales(1) = "L010"
DSLoc(3).locales(2) = "L011"
DSLoc(3).locales(3) = "L012"
DSLoc(3).locales(4) = "L020"
DSLoc(3).locales(5) = "L030"

DSLoc(4).Dep = "INT"
DSLoc(4).Sdep = "CORD"
DSLoc(4).locales(1) = "L040"
DSLoc(4).SLocales(1) = 0
DSLoc(4).locales(2) = "L190"
DSLoc(4).SLocales(2) = 0

DSLoc(5).Dep = "INT"
DSLoc(5).Sdep = "IGUA"
DSLoc(5).locales(1) = "L110"
DSLoc(5).SLocales(1) = 0

DSLoc(6).Dep = "INT"
DSLoc(6).Sdep = "MEND"
DSLoc(6).locales(1) = "L100"
DSLoc(6).SLocales(1) = 0
DSLoc(6).locales(2) = "L160"
DSLoc(6).SLocales(2) = 0

DSLoc(7).Dep = "INT"
DSLoc(7).Sdep = "MDPL"
DSLoc(7).locales(1) = "L150"
DSLoc(7).SLocales(1) = 0

DSLoc(8).Dep = "INT"
DSLoc(8).Sdep = "BARI"
DSLoc(8).locales(1) = "L180"
DSLoc(8).SLocales(1) = 0


DSLoc(11).Dep = "IFL"
DSLoc(11).Sdep = "IFL"
DSLoc(11).locales(1) = ""
DSLoc(11).SLocales(1) = 0

DSLoc(12).Dep = "IFL"
DSLoc(12).Sdep = "AME"
DSLoc(12).locales(1) = ""
DSLoc(12).SLocales(1) = 0



End Sub

Private Function MePerfilUser() As Integer
Dim sql$, SQLJ$
Dim rs As Recordset, RSJefe As Recordset
Dim NU As String
Dim PU As String

On Error GoTo ErrDatos:
    
    sql$ = ""
    sql$ = sql$ & "Select PERFIL,nivel "
    sql$ = sql$ & " FROM estadis.PERFILES "
    sql$ = sql$ & " WHERE usuario = '" & Aplicacion.username & "'"

    If Aplicacion.ObtenerRsDAO(sql$, rs) Then
        If Aplicacion.CantReg(rs) > 0 Then
            Aplicacion.Perfil = rs!Perfil
            Aplicacion.Nivel = rs!Nivel
        Else
            Aplicacion.Perfil = ""
            Aplicacion.Nivel = -1
        End If
        Aplicacion.CerrarDAO rs
    Else
        MsgBox "Existe algún problema con su Nivel de Acceso. Se le otorgará el mínimo posible.", vbExclamation + vbOKOnly, "ATENCION"
        Aplicacion.Perfil = ""
        Aplicacion.Nivel = -1

    End If

ErrDatos:
    Exit Function
    
End Function
