VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLProdMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private l_cod_prod As Long
Private s_descrip As String


Public Property Get Cod_prod() As Long
Cod_prod = l_cod_prod
End Property

Public Property Let Cod_prod(ByVal vNewValue As Long)
l_cod_prod = vNewValue
End Property

Public Property Get Descrip() As String
Descrip = s_descrip
End Property

Public Property Let Descrip(ByVal vNewValue As String)
    s_descrip = vNewValue
End Property
