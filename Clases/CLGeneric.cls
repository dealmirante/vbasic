VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private s_Identif As String
Private si_Porc As Single
Private s_Depn  As String

Private s_Tipo As String
Private s_Grupo As String
Private s_Local As String

Private si_PorcGA As Single
Private si_PorcGB As Single
Private si_PorcGC As Single

Private si_VALOR_1 As Single
Private si_VALOR_2 As Single
Private si_VALOR_3 As Single
Private si_VALOR_4 As Single
Private si_VALOR_5 As Single

Private s_tipo_comi As String
Public Property Get Identif() As Variant
Identif = s_Identif
End Property

Public Property Let Identif(ByVal vNewValue As Variant)
s_Identif = vNewValue
End Property

Public Property Get Porc() As Variant
Porc = si_Porc
End Property

Public Property Let Porc(ByVal vNewValue As Variant)
si_Porc = vNewValue
End Property

Public Property Get PorcGA() As Variant
PorcGA = si_PorcGA
End Property

Public Property Let PorcGA(ByVal vNewValue As Variant)
si_PorcGA = vNewValue
End Property

Public Property Get PorcGB() As Variant
PorcGB = si_PorcGB
End Property

Public Property Let PorcGB(ByVal vNewValue As Variant)
si_PorcGB = vNewValue
End Property

Public Property Get PorcGC() As Variant
PorcGC = si_PorcGC
End Property

Public Property Let PorcGC(ByVal vNewValue As Variant)
si_PorcGC = vNewValue
End Property

Public Property Get depn() As Variant
depn = s_Depn
End Property

Public Property Let depn(ByVal vNewValue As Variant)
s_Depn = vNewValue
End Property

Public Property Get Grupo() As Variant
Grupo = s_Grupo
End Property

Public Property Let Grupo(ByVal vNewValue As Variant)
s_Grupo = vNewValue
End Property

Public Property Get Tipo() As Variant
Tipo = s_Tipo
End Property

Public Property Let Tipo(ByVal vNewValue As Variant)
s_Tipo = vNewValue
End Property
Public Property Get Locale() As Variant
Locale = s_Local
End Property

Public Property Let Locale(ByVal vNewValue As Variant)
s_Local = vNewValue
End Property

Public Property Get VALOR_1() As Variant
VALOR_1 = si_VALOR_1
End Property

Public Property Let VALOR_1(ByVal vNewValue As Variant)
si_VALOR_1 = vNewValue
End Property

Public Property Get VALOR_2() As Variant
VALOR_2 = si_VALOR_2
End Property

Public Property Let VALOR_2(ByVal vNewValue As Variant)
si_VALOR_2 = vNewValue
End Property

Public Property Get VALOR_3() As Variant
VALOR_3 = si_VALOR_3
End Property

Public Property Let VALOR_3(ByVal vNewValue As Variant)
si_VALOR_3 = vNewValue
End Property

Public Property Get VALOR_4() As Variant
VALOR_4 = si_VALOR_4
End Property

Public Property Let VALOR_4(ByVal vNewValue As Variant)
si_VALOR_4 = vNewValue
End Property

Public Property Get VALOR_5() As Variant
VALOR_5 = si_VALOR_5
End Property

Public Property Let VALOR_5(ByVal vNewValue As Variant)
si_VALOR_5 = vNewValue
End Property

