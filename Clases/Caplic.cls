VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sUsername As String
Private m_sPassword As String

'Variables de acceso a datos
Private i_Nivel As Integer
Private s_Perfil As String
Private i_Anio As Integer

'Private m_rAmbiente As rdoEnvironment

'DAO
Private m_dbDAO As Database
Private m_Space As Workspace

'Defino el datasource name
Private m_sDsn As String

'Vaiabbles de control auxiliares
Private m_sCaptionForm As String
Private m_sCaptionPross As String


Public Function CantReg(ByVal rs As Recordset)

If rs.RecordCount > 0 Then
    rs.MoveLast
    CantReg = rs.RecordCount
    rs.MoveFirst
Else
    CantReg = 0
End If
End Function

Public Sub CerrarDAO(rs As Recordset)
    rs.Close
End Sub


Public Function ObtenerDyDAO(sql As String, ByRef rs As Recordset)
On Error GoTo ErrorDB:

    Set rs = m_dbDAO.OpenRecordset(sql, dbOpenDynaset, dbSQLPassThrough)
    ObtenerDyDAO = True
Exit Function
ErrorDB:
        Call ManejadorDeErrores
        ObtenerDyDAO = False

End Function

Public Property Let Perfil(ByVal vNewValue As String)
    s_Perfil = vNewValue
End Property

Public Sub TerminarConErrorTrans()
'    m_rConexion.RollbackTrans
'm_dbDAO.RollbackTrans
m_Space.Rollback
End Sub

Public Sub TerminarConExitoTrans()
' m_dbDAO.Execute ("commit"), dbSQLPassThrough
 m_Space.CommitTrans
End Sub

Public Property Get username() As String
    username = m_sUsername
End Property

Public Property Let username(vNewValue As String)
    m_sUsername = vNewValue
End Property

Public Property Get password()
    password = m_sPassword
End Property

Public Property Let password(vNewValue)
    m_sPassword = vNewValue
End Property

Public Function conectarDAO() As Boolean
Dim sConexion As String
 
 On Error GoTo ErrConect:

Aplicacion.SeteoProceso ""

';SERVER=TNS:INTER
    'sConexion = "ODBC;UID=" & m_sUsername & ";PWD=" & m_sPassword & ";DATABASE=dutypaid;"
    sConexion = "ODBC;UID=" & m_sUsername & ";PWD=" & m_sPassword & ";"
    Set m_Space = Workspaces(0)
    Set m_dbDAO = m_Space.OpenDatabase("inter", False, False, sConexion)
 
    conectarDAO = True
    Aplicacion.SeteoFin
Exit Function

ErrConect:
        Aplicacion.SeteoFin
        MsgBox "Error en Conexión", vbCritical + vbOK
        conectarDAO = False

End Function

Public Function conectar() As Boolean
Dim sConexion As String

'Función de conexión a la base de datos

 On Error GoTo ErrConect:

    sConexion = "ODBC;DSN=INTER;UID=" & m_sUsername & ";PWD=" & m_sPassword
    
'    Set m_rAmbiente = rdoEnvironments(0)
    
'    m_rAmbiente.CursorDriver = rdUseOdbc
'   m_rAmbiente.CursorDriver = rdUseServer
'   m_rAmbiente.CursorDriver = rdUseIfNeeded

    
'    Set m_rConexion = m_rAmbiente.OpenConnection("", rdDriverNoPrompt, False, sConexion)
        
'    conectar = True
    
Exit Function
ErrConect:
        MsgBox "Error en Conexión", vbCritical + vbOK
        conectar = False
End Function

'Public Sub obtenerRsSinc(sql As String, ByRef RS As rdoResultset)
'Devuelve un resultset ejecutándose en forma sincrónica
'Solo funciona ReadOnly con el driver Visigenic

'Set rs = m_rConexion.OpenResultset(sql, rdOpenStatic, rdConcurRowver)

'End Sub



Public Sub EjecutarSinc(sql As String)
'Se ejecuta la sentencia en forma sinc
'Es para sentencias de acción
'    m_rConexion.Execute sql, rdExecDirect
End Sub

Public Function EjecutarDAO(sql As String)

On Error GoTo ErrorDB:
        
    m_dbDAO.Execute sql, dbSQLPassThrough
    
    EjecutarDAO = True

Exit Function
ErrorDB:
        Call ManejadorDeErrores
        EjecutarDAO = False
End Function

Public Function ObtenerRsDAO(sql As String, ByRef rs As Recordset)
    
On Error GoTo ErrorDB:

    Set rs = m_dbDAO.OpenRecordset(sql, dbOpenSnapshot, dbSQLPassThrough)
    ObtenerRsDAO = True
Exit Function
ErrorDB:
        Call ManejadorDeErrores
        ObtenerRsDAO = False
End Function

Public Property Get Ndatabase() As Database
'    Ndatabase = m_dbDAO
End Property

Public Property Let Ndatabase(vNewValue As Database)
'    m_dbDAO = vNewValue
End Property

Public Property Get NdbRDO()
'    NdbRDO = m_rConexion
End Property

Public Property Let NdbRDO(vNewValue)
'    m_rConexion = vNewValue
End Property







Public Sub ComienzoTrans()
'    m_rConexion.BeginTrans
  m_Space.BeginTrans
End Sub


'Public Sub CerrarRS(RS As rdoResultset)

'    RS.Close
    
'End Sub

Public Sub ManejadorDeErrores()
Dim clError As Error


'MsgBox Errors.Count
'For Each clError In DBEngine.Errors
'MsgBox clError.Number & " " & clError.Description
'Next
Set clError = Errors.item(0)

Select Case clError.Number
    Case 1
        MsgBox "Clave Duplicada", vbCritical + vbOKOnly, "ERROR"
    Case 54
        MsgBox "Recurso Ocupado ", vbCritical + vbOKOnly, "ERROR"
    Case 2292
        MsgBox "Existen Datos dependientes", vbCritical + vbOKOnly, "ERROR"
    Case 1407
        MsgBox "Falta Ingresar información", vbCritical + vbOKOnly, "ERROR"
    Case Else
        MsgBox clError.Number & " " & clError.Description, vbCritical + vbOKOnly, "ERROR"
End Select

End Sub

Public Property Get CaptionForm() As String
CaptionForm = m_sCaptionForm
End Property

Public Property Let CaptionForm(vNewValue As String)
m_sCaptionForm = vNewValue
End Property

Public Property Get CaptionPross() As String
CaptionPross = m_sCaptionPross
End Property

Public Property Let CaptionPross(vNewValue As String)
m_sCaptionPross = vNewValue
End Property

Public Function SeteoProceso(caption As String) As String
    
    Screen.MousePointer = vbHourglass
    CaptionForm = caption
    SeteoProceso = CaptionPross
    
End Function
Public Function SeteoFin() As String
    
    Screen.MousePointer = vbDefault
    SeteoFin = CaptionForm
    CaptionForm = ""
    
End Function





Public Property Get Perfil() As String
Perfil = s_Perfil
End Property




Public Property Get Nivel() As Integer
Nivel = i_Nivel
End Property

Public Property Let Nivel(ByVal vNewValue As Integer)
i_Nivel = vNewValue
End Property

Public Property Get anio() As Variant
anio = i_Anio
End Property

Public Property Let anio(ByVal vNewValue As Variant)
i_Anio = vNewValue
End Property
