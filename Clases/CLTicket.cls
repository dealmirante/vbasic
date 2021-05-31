VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private i_Nro_Ticket As Integer
Private s_Cod_Sdep As String
Private s_Cod_Local As String
Private s_Cod_Sloc As String
Private i_Cod_Caja As Integer
Private s_Grupo_Venta As String
Private i_Cod_Cia_Aerea As Integer
Private s_Descrip As String
Private i_Cod_Vuelo As Integer
Private i_Nro_Pax As Integer
Private l_Cod_Cajero As Long
Private l_Legajo As Long
Private s_Apellido As String
Private s_Nombre As String
Private s_Credit As String

Public Property Get Cod_Sdep() As Variant
Cod_Sdep = s_Cod_Sdep
End Property
Public Property Get Descrip() As Variant
Descrip = s_Descrip
End Property


Public Property Get Apellido() As Variant
Apellido = s_Apellido
End Property


Public Property Get Credit() As Variant
Credit = s_Credit
End Property
Public Property Get Nombre() As Variant
Nombre = s_Nombre
End Property
Public Property Get Grupo_Venta() As Variant
Grupo_Venta = s_Grupo_Venta
End Property
Public Property Get Cod_Local() As Variant
Cod_Local = s_Cod_Local
End Property

Public Property Get Cod_Sloc() As Variant
Cod_Sloc = s_Cod_Sloc
End Property
Public Property Let Descrip(ByVal vNewValue As Variant)
s_Descrip = vNewValue
End Property

Public Property Let Credit(ByVal vNewValue As Variant)
s_Credit = vNewValue
End Property
Public Property Let Apellido(ByVal vNewValue As Variant)
s_Apellido = vNewValue
End Property


Public Property Let Nombre(ByVal vNewValue As Variant)
s_Nombre = vNewValue
End Property
Public Property Let Grupo_Venta(ByVal vNewValue As Variant)
s_Grupo_Venta = vNewValue
End Property
Public Property Let Cod_Local(ByVal vNewValue As Variant)
s_Cod_Local = vNewValue
End Property

Public Property Let Cod_Sloc(ByVal vNewValue As Variant)
s_Cod_Sloc = vNewValue
End Property
Public Property Let Cod_Sdep(ByVal vNewValue As Variant)
s_Cod_Sdep = vNewValue
End Property
Public Property Get Nro_Ticket() As Variant
Nro_Ticket = i_Nro_Ticket
End Property

Public Property Get Cod_Caja() As Variant
Cod_Caja = i_Cod_Caja
End Property

Public Property Get Nro_Pax() As Variant
Nro_Pax = i_Nro_Pax
End Property
Public Property Get Cod_Vuelo() As Variant
Cod_Vuelo = i_Cod_Vuelo
End Property
Public Property Get Cod_Cia_Aerea() As Variant
Cod_Cia_Aerea = i_Cod_Cia_Aerea
End Property
Public Property Get Legajo() As Variant
Legajo = l_Legajo
End Property


Public Property Get Cod_Cajero() As Variant
Cod_Cajero = l_Cod_Cajero
End Property
Public Property Let Cod_Cajero(ByVal vNewValue As Variant)
l_Cod_Cajero = vNewValue
End Property
Public Property Let Nro_Ticket(ByVal vNewValue As Variant)
i_Nro_Ticket = vNewValue
End Property

Public Property Let Cod_Caja(ByVal vNewValue As Variant)
i_Cod_Caja = vNewValue
End Property

Public Property Let Nro_Pax(ByVal vNewValue As Variant)
i_Nro_Pax = vNewValue
End Property
Public Property Let Cod_Vuelo(ByVal vNewValue As Variant)
i_Cod_Vuelo = vNewValue
End Property
Public Property Let Cod_Cia_Aerea(ByVal vNewValue As Variant)
i_Cod_Cia_Aerea = vNewValue
End Property
