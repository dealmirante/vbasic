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
Global Const INTE = "INTE"
Global Const Total = "TOTAL"
Global Const EZEAL = "EZEAL"
Global Const EZEAS = "EZEAS"

Global Const BARI = "BARI"
Global Const CORD = "CORD"
Global Const MDPL = "MDPL"
Global Const MEND = "MEND"

Global Const TotalINTA = 11

Global Const NroLocINTA = 9
Global Const NroLocINTB = 5
Global Const NroLocAERO = 3
Global Const NroLocINTE = 6

Global Const part = "P"
Global Const FULL = "F"


'dias de semana
Global Const DOMINGO = 1
Global Const SABADO = 7

Global Const ESQUEMA = "estadis."

'Tipo de datos para aparear los locales de cada dependencia

Type T_Locales
    Dep As String
    Sdep As String
    locales(1 To 15) As String
    SLocales(1 To 15) As Integer
End Type

Global DSLoc(1 To 15) As T_Locales

Type T_Col
    sLocal As String
    iCol As Integer
End Type

Type T_ConcProv
    codigo As String
    sec As Integer
End Type
