VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLVersus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private d_fch_vuelo As Variant
Private i_cod_cia_aerea As Integer
Private i_cod_vuelo As Integer
Private s_cod_depn As String
Private s_cod_sdep As String
Private s_Tipo As String
Private i_pax_at As Integer
Private i_pax_vol As Integer
Private s_des_cia As String
Private i_cod_nuevo_vuelo As Integer
Private i_nuevo_pax_at As Integer

Public Property Get fch_vuelo() As Variant
fch_vuelo = d_fch_vuelo
End Property

Public Property Let fch_vuelo(ByVal vNewValue As Variant)
d_fch_vuelo = vNewValue
End Property

Public Property Get cod_cia_aerea() As Variant
cod_cia_aerea = i_cod_cia_aerea
End Property

Public Function insert_pax_at_vs_vol() As Boolean

Dim sql As String
sql = "INSERT INTO "
sql = sql & " estadis.pax_at_vs_vol ("
sql = sql & " fch_vuelo,"
sql = sql & " cod_cia_vuelo,"
sql = sql & " cod_vuelo,"
sql = sql & " cod_depn,"
sql = sql & " cod_sdep,"
sql = sql & " tipo,"
sql = sql & " pax_at,"
sql = sql & " pax_vol ) "
sql = sql & " VALUES ("
sql = sql & "'" & fch_vuelo & "',"
sql = sql & "'" & cod_cia_vuelo & "',"
sql = sql & "'" & cod_vuelo & "',"
sql = sql & "'" & cod_depn & "',"
sql = sql & "'" & cod_sdep & "',"
sql = sql & "'" & tipo & "',"
sql = sql & "'" & pax_at & "',"
sql = sql & "'" & pax_vol & "')"

Insert_insert_pax_at_vs_vol = Aplicacion.EjecutarDAO(sql)

End Function

Public Function update_pax_at_vs_vol() As Boolean

Dim sql As String

sql = "UPDATE "
sql = sql & " estadis.pax_at_vs_vol SET "
sql = sql & " pax_at = " & pax_at
sql = sql & " WHERE fch_vuelo =" & func_ToDate(fch_vuelo)
sql = sql & " AND cod_cia_aerea =" & cod_cia_aerea
sql = sql & " AND cod_vuelo =" & cod_vuelo
sql = sql & " AND cod_depn =" & "'" & cod_depn & "'"
sql = sql & " AND cod_sdep =" & "'" & cod_sdep & "'"

update_pax_at_vs_vol = Aplicacion.EjecutarDAO(sql)

'Actualizo los datos del nuevo vuelo

sql = "UPDATE "
sql = sql & " estadis.pax_at_vs_vol SET "
sql = sql & " pax_at = " & nuevo_pax_at
sql = sql & " WHERE fch_vuelo =" & func_ToDate(fch_vuelo)
sql = sql & " AND cod_cia_aerea =" & cod_cia_aerea
sql = sql & " AND cod_vuelo =" & cod_nuevo_vuelo
sql = sql & " AND cod_depn =" & "'" & cod_depn & "'"
sql = sql & " AND cod_sdep =" & "'" & cod_sdep & "'"

update_pax_at_vs_vol = Aplicacion.EjecutarDAO(sql)


End Function
Public Property Let cod_cia_aerea(ByVal vNewValue As Variant)
i_cod_cia_aerea = vNewValue
End Property

Public Property Get cod_vuelo() As Variant
cod_vuelo = i_cod_vuelo
End Property


Public Property Let cod_vuelo(ByVal vNewValue As Variant)
i_cod_vuelo = vNewValue
End Property

Public Property Get cod_depn() As Variant
cod_depn = s_cod_depn
End Property

Public Property Let cod_depn(ByVal vNewValue As Variant)
s_cod_depn = vNewValue
End Property

Public Property Get cod_sdep() As Variant
cod_sdep = s_cod_sdep
End Property

Public Property Let cod_sdep(ByVal vNewValue As Variant)
s_cod_sdep = vNewValue
End Property

Public Property Get cod_tipo() As Variant
cod_tipo = s_cod_tipo
End Property

Public Property Let cod_tipo(ByVal vNewValue As Variant)
s_cod_tipo = vNewValue
End Property

Public Property Get pax_at() As Variant
pax_at = i_pax_at
End Property

Public Property Let pax_at(ByVal vNewValue As Variant)
i_pax_at = vNewValue
End Property

Public Property Get pax_vol() As Variant
pax_vol = i_pax_vol
End Property

Public Property Let pax_vol(ByVal vNewValue As Variant)
i_pax_vol = vNewValue
End Property

Public Property Get des_cia() As Variant
des_cia = s_des_cia
End Property

Public Property Let des_cia(ByVal vNewValue As Variant)
s_des_cia = vNewValue
End Property

Public Property Get cod_nuevo_vuelo() As Variant
cod_nuevo_vuelo = i_cod_nuevo_vuelo
End Property

Public Property Let cod_nuevo_vuelo(ByVal vNewValue As Variant)
i_cod_nuevo_vuelo = vNewValue
End Property

Public Property Get nuevo_pax_at() As Variant
nuevo_pax_at = i_nuevo_pax_at
End Property

Public Property Let nuevo_pax_at(ByVal vNewValue As Variant)
i_nuevo_pax_at = vNewValue
End Property
