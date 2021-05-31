Attribute VB_Name = "mPrincipal"
Option Explicit

Public Sub Main()
Const iMaxIntentos = 3
Dim iIntentos As Integer
Dim bSalir As Boolean
Dim bConecto As Boolean, bConectoP As Boolean
Dim Ruta As String
Dim clave As String
Dim Seccion As String
Dim cant As Integer
Dim StrRet As String
Dim defol As String

Ruta = "win.ini"
Seccion = "Intl"
clave = "sDecimal"
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

Call GetPrivateProfileString(Seccion, clave, defol, StrRet, cant, Ruta)

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

    If FrmLogon.PedirUsuario Then 'pedir usuario muestra la pantalla de log y si el usuario continúa

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

Select Case Aplicacion.Nivel
    Case 0
        frmMDI.mnuGral(0).Visible = True
        frmMDI.mnuGral(1).Visible = True
        frmMDI.mnuGral(4).Visible = True
        frmMDI.mnuGral(5).Visible = True
        frmMDI.mnuGral(6).Visible = True
    Case 1
        frmMDI.mnuGral(0).Visible = True
        frmMDI.mnuGral(1).Visible = True
        frmMDI.mnuGral(4).Visible = True
        frmMDI.mnuGral(5).Visible = True
'        frmMDI.mnuGral(6).Visible = True
'        frmMDI.mnuCom(0).Visible = False
'        frmMDI.mnuCom(1).Visible = False
'        frmMDI.mnuCom(2).Visible = False
'        frmMDI.mnuCom(3).Visible = False
    Case 2
        frmMDI.mnuGral(0).Visible = True
        frmMDI.mnuGral(1).Visible = True
        frmMDI.mnuGral(4).Visible = True
        frmMDI.mnuGral(5).Visible = True
'        frmMDI.mnuCom(0).Visible = False
'        frmMDI.mnuCom(1).Visible = False
'        frmMDI.mnuCom(2).Visible = False
'        frmMDI.mnuCom(3).Visible = False
        frmMDI.mnuCom(4).Visible = False
        frmMDI.mnuCap(0).Visible = False
        frmMDI.mnuCap(1).Visible = False
    Case 10
        frmMDI.mnuGral(0).Visible = True
        frmMDI.mnuGral(1).Visible = True
        frmMDI.mnuGral(4).Visible = True
        frmMDI.mnuGral(5).Visible = True
'        frmMDI.mnuCom(0).Visible = False
'        frmMDI.mnuCom(1).Visible = False
'        frmMDI.mnuCom(2).Visible = True
'        frmMDI.mnuCom(3).Visible = True
        frmMDI.mnuCom(4).Visible = False
End Select
End Sub


Private Sub MeLlenarDSLocal()
DSLoc(1).Dep = "AEP"
DSLoc(1).Sdep = "AEP"
DSLoc(1).Locales(1) = "L090"
DSLoc(1).SLocales(1) = 0
DSLoc(1).Locales(2) = "L140"
DSLoc(1).SLocales(2) = 0
DSLoc(1).Locales(3) = "L141"
DSLoc(1).SLocales(3) = 1


DSLoc(2).Dep = "EZE"
DSLoc(2).Sdep = "INTA"
DSLoc(2).Locales(1) = "L050"
DSLoc(2).SLocales(1) = 0
DSLoc(2).Locales(2) = "L060"
DSLoc(2).SLocales(2) = 0
DSLoc(2).Locales(3) = "L070"
DSLoc(2).SLocales(3) = 0
DSLoc(2).Locales(4) = "L071"
DSLoc(2).SLocales(4) = 1
DSLoc(2).Locales(5) = "L080"
DSLoc(2).SLocales(5) = 0
DSLoc(2).Locales(6) = "L081"
DSLoc(2).SLocales(6) = 1
DSLoc(2).Locales(7) = "L082"
DSLoc(2).SLocales(7) = 2


DSLoc(3).Dep = "EZE"
DSLoc(3).Sdep = "INTB"
DSLoc(3).Locales(1) = "L010"
DSLoc(3).Locales(2) = "L011"
DSLoc(3).Locales(3) = "L020"
DSLoc(3).Locales(4) = "L030"


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
