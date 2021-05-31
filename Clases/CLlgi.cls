VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLlgi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private s_fch As String
Private s_Depn As String
Private s_Sdep As String
Private s_Local As String
Private i_sloc As Integer
Private i_imp As Double
Private l_cant As Long
Private l_Pax As Long
Private i_Dato As Double

Public col_prnLGI As Collection
Public Property Get fch() As Variant
fch = s_fch
End Property

Public Property Let fch(ByVal vNewValue As Variant)
s_fch = vNewValue
End Property

Public Property Get Depn() As Variant
Depn = s_Depn
End Property

Public Property Let Depn(ByVal vNewValue As Variant)
s_Depn = vNewValue
End Property


Public Property Get Sdep() As Variant
Sdep = s_Sdep
End Property

Public Property Let Sdep(ByVal vNewValue As Variant)
s_Sdep = vNewValue
End Property

Public Property Get Locale() As Variant
Locale = s_Local
End Property

Public Property Let Locale(ByVal vNewValue As Variant)
s_Local = vNewValue
End Property


Public Property Get Importe() As Double
Importe = i_imp
End Property

Public Property Let Importe(ByVal vNewValue As Double)
i_imp = vNewValue
End Property

Public Property Get Ticket() As Long
Ticket = l_cant
End Property

Public Property Let Ticket(ByVal vNewValue As Long)
l_cant = vNewValue
End Property

Public Property Get Pax() As Long
Pax = l_Pax
End Property

Public Property Let Pax(ByVal vNewValue As Long)
l_Pax = vNewValue
End Property

Public Property Get DatoGral() As Double
DatoGral = i_Dato
End Property

Public Property Let DatoGral(ByVal vNewValue As Double)
i_Dato = vNewValue
End Property

