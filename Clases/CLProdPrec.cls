VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLProdPrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private l_CodProd As Long
Private si_Precio As Single
Private s_CodProv As String
Private s_CodRubr As String
Private i_sec As Integer
Private i_Tipo As Integer
Private i_Objetivo As Single

Private s_fdesde As String
Private s_fhasta As String

Public Property Get codProd() As Variant
codProd = l_CodProd
End Property
Public Property Let codProd(ByVal vNewValue As Variant)
l_CodProd = vNewValue
End Property
Public Property Get Precio() As Variant
Precio = si_Precio
End Property
Public Property Let Precio(ByVal vNewValue As Variant)
si_Precio = vNewValue
End Property
Public Property Get CodProv() As Variant
CodProv = s_CodProv
End Property
Public Property Let CodProv(ByVal vNewValue As Variant)
s_CodProv = vNewValue
End Property
Public Property Get sec() As Variant
sec = i_sec
End Property
Public Property Let sec(ByVal vNewValue As Variant)
i_sec = vNewValue
End Property
Public Property Get tipo() As Variant
tipo = i_Tipo
End Property
Public Property Let tipo(ByVal vNewValue As Variant)
i_Tipo = vNewValue
End Property
Public Property Get Objetivo() As Variant
Objetivo = i_Objetivo
End Property
Public Property Let Objetivo(ByVal vNewValue As Variant)
i_Objetivo = vNewValue
End Property

Public Property Get fdesde() As Variant
fdesde = s_fdesde
End Property

Public Property Let fdesde(ByVal vNewValue As Variant)
s_fdesde = vNewValue
End Property

Public Property Get fhasta() As Variant
fhasta = s_fhasta
End Property

Public Property Let fhasta(ByVal vNewValue As Variant)
s_fhasta = vNewValue
End Property

Public Property Get CodRubr() As Variant
CodRubr = s_CodRubr
End Property

Public Property Let CodRubr(ByVal vNewValue As Variant)
s_CodRubr = vNewValue
End Property
