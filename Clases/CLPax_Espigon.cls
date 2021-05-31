VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLPax_Espigon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private d_fch_vta As Date
Private i_cod_cia_aerea As Integer
Private i_cod_vuelo As Integer
Private s_cod_depn As String
Private s_cod_sdep As String
Private i_nacionalidad As Integer
Private i_cant_tickets As Integer
Private i_cant_pax As Integer


Public Property Get fch_vta() As Variant
d_fch_vta = fch_vta
End Property

Public Property Let fch_vta(ByVal vNewValue As Variant)
d_fch_vta = vNewValue
End Property

Public Property Get cod_cia_aerea() As Variant
i_cod_cia_aerea = cod_cia_aerea
End Property

Public Property Let cod_cia_aerea(ByVal vNewValue As Variant)
i_cod_cia_aerea = vNewValue
End Property

Public Property Get cod_vuelo() As Variant
i_cod_vuelo = cod_vuelo
End Property

Public Property Let cod_vuelo(ByVal vNewValue As Variant)
i_cod_vuelo = vNewValue
End Property

Public Property Get cod_depn() As Variant
s_cod_depn = cod_depn
End Property

Public Property Let cod_depn(ByVal vNewValue As Variant)
s_cod_depn = vNewValue
End Property

Public Property Get cod_sdep() As Variant
s_cod_sdep = cod_sdep
End Property

Public Property Let cod_sdep(ByVal vNewValue As Variant)
s_cod_sdep = vNewValue
End Property

Public Function Insert_Pax_Espigon() As Boolean
Dim sql As String
sql = "INSERT INTO "
sql = sql & " estadis.pax_espigon ("
sql = sql & " fch_vta,"
sql = sql & " cod_cia_vuelo,"
sql = sql & " cod_vuelo,"
sql = sql & " cod_depn,"
sql = sql & " cod_sdep,"
sql = sql & " nacionalidad,"
sql = sql & " cant_tickets,"
sql = sql & " cant_pax ) "
sql = sql & " VALUES ("
sql = sql & "'" & fch_vta & "',"
sql = sql & "'" & cod_cia_vuelo & "',"
sql = sql & "'" & cod_vuelo & "',"
sql = sql & "'" & cod_depn & "',"
sql = sql & "'" & cod_sdep & "',"
sql = sql & "'" & nacionalidad & "',"
sql = sql & "'" & cant_tickets & "',"
sql = sql & "'" & cant_pax & "')"

Insert_Pax_Espigon = Aplicacion.EjecutarDAO(sql)

End Function

Public Property Get nacionalidad() As Variant
i_nacionalidad = nacionalidad
End Property

Public Property Let nacionalidad(ByVal vNewValue As Variant)
i_nacionalidad = vNewValue
End Property

Public Property Get cant_tickets() As Variant
i_cant_tickets = cant_tickets
End Property

Public Property Let cant_tickets(ByVal vNewValue As Variant)
i_cant_tickets = vNewValue
End Property

Public Property Get cant_pax() As Variant
cant_pax = i_cant_pax
End Property

Public Property Let cant_pax(ByVal vNewValue As Variant)
i_cant_pax = vNewValue
End Property
Public Function update_pax_espigon() As Boolean

Dim sql As String

sql = "UPDATE "
sql = sql & " estadis.pax_espigon SET "
sql = sql & " fch_vta =" & "'" & fch_vta & "',"
sql = sql & " cod_cia_aerea =" & "'" & cod_cia_aerea & "',"
sql = sql & " cod_vuelo =" & "'" & cod_vuelo & "',"
sql = sql & " cod_depn =" & "'" & cod_depn & "',"
sql = sql & " cod_sdep =" & "'" & cod_sdep & "',"
sql = sql & " nacionalidad =" & "'" & nacionalidad & "',"
sql = sql & " cant_tickets =" & "'" & cant_tickets & "',"
sql = sql & " cant_pax =" & "'" & cant_pax & "'"

update_pax_espigon = Aplicacion.EjecutarDAO(sql)

End Function


