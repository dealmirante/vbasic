Attribute VB_Name = "General"
Global g_cursor As String

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Global Aplicacion As CAplicacion

Global RutaFotos As String
Global RutaReportes As String
Global CantAnio As Integer

'Constantes de Formatos
Global Const MedioAncho = 1

Global Const FTOFECHA = "dd-mm-yyyy"
Global Const FTOFECHAHORA = "dd-mm-yyyy hh:mm:ss"


'Constantes de valores
Global Const NIVEL1 = 1

Global Const EZEA = "EZEA"
Global Const EZEB = "EZEB"
Global Const AERO = "AERO"
Global Const total = "TOTAL"

Global Const TotalINTA = 9

Global Const NroLocINTA = 7
Global Const NroLocINTB = 4
Global Const NroLocAERO = 3

'dias de semana
Global Const DOMINGO = 1
Global Const SABADO = 7

Global Const ESQUEMA = "estadis."

'Tipo de datos para aparear los locales de cada dependencia

Type T_Locales
    Dep As String
    Sdep As String
    Locales(1 To 12) As String
    SLocales(1 To 12) As Integer
End Type

Global DSLoc(1 To 10) As T_Locales

Type T_Col
    sLocal As String
    iCol As Integer
End Type

Type T_ConcProv
    codigo As String
    Sec As Integer
End Type
