VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLVolados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private d_fch_vuelo As String
Private i_cod_cia_aerea As Integer
Private i_cod_vuelo As Integer
Private s_cod_depn As String
Private s_cod_sdep As String
Private i_hora_vuelo As Integer
Private i_nacionalidad As Integer
Private i_cantE As Integer
Private i_cantN As Integer
Private i_cantT As Integer
Private s_Tipo As String
Private s_Local As String
Private s_Destino As String
Public Property Get fch_vuelo() As String
fch_vuelo = d_fch_vuelo
End Property
Public Property Let fch_vuelo(ByVal vNewValue As String)
d_fch_vuelo = vNewValue
End Property
Public Property Get cod_cia_aerea() As Integer
cod_cia_aerea = i_cod_cia_aerea
End Property
Public Property Let cod_cia_aerea(ByVal vNewValue As Integer)
i_cod_cia_aerea = vNewValue
End Property
Public Property Get cod_vuelo() As Integer
cod_vuelo = i_cod_vuelo
End Property
Public Property Let cod_vuelo(ByVal vNewValue As Integer)
i_cod_vuelo = vNewValue
End Property
Public Property Get cod_depn() As String
cod_depn = s_cod_depn
End Property

Public Property Let cod_depn(ByVal vNewValue As String)
s_cod_depn = vNewValue
End Property
Public Property Get Destino() As String
Destino = s_Destino
End Property

Public Property Let Destino(ByVal vNewValue As String)
s_Destino = vNewValue
End Property

Public Property Get cod_sdep() As String
cod_sdep = s_cod_sdep
End Property

Public Property Let cod_sdep(ByVal vNewValue As String)
s_cod_sdep = vNewValue
End Property

Public Property Get hora_vuelo() As Integer
hora_vuelo = i_hora_vuelo
End Property

Public Property Let hora_vuelo(ByVal vNewValue As Integer)
i_hora_vuelo = vNewValue
End Property

Public Function Insert_Pax_Volados() As Boolean

Dim sql As String

'Inserta un registro con cantidad de pax extranjeros
sql = "INSERT INTO "
sql = sql & " Estadis.Pax_Volados ("
sql = sql & " fch_vuelo,"
sql = sql & " cod_cia_aerea,"
sql = sql & " cod_vuelo,"
sql = sql & " cod_depn,"
sql = sql & " cod_sdep,"
sql = sql & " hora_vuelo,"
sql = sql & " nacionalidad,"
sql = sql & " cantidad,"
sql = sql & " tipo, "
sql = sql & " local,"
sql = sql & " transito ) "
sql = sql & " VALUES ("
sql = sql & func_ToDate(fch_vuelo) & ","
sql = sql & "" & cod_cia_aerea & ","
sql = sql & "" & cod_vuelo & ","
sql = sql & "'" & cod_depn & "',"
sql = sql & "'" & cod_sdep & "',"
sql = sql & "" & hora_vuelo & ","
sql = sql & "'E',"
sql = sql & "" & (cantE + cantT) & ","
sql = sql & "'" & Tipo & "',"
sql = sql & "'" & locales & "',"
sql = sql & cantT & ")"

If Aplicacion.EjecutarDAO(sql) Then

'Inserta un registro para cantidad de pax nacionales

sql = "INSERT INTO "
sql = sql & " estadis.Pax_Volados ("
sql = sql & " fch_vuelo, "
sql = sql & " cod_cia_aerea,"
sql = sql & " cod_vuelo,"
sql = sql & " cod_depn,"
sql = sql & " cod_sdep,"
sql = sql & " hora_vuelo,"
sql = sql & " nacionalidad,"
sql = sql & " cantidad,"
sql = sql & " tipo, "
sql = sql & " local, "
sql = sql & " transito ) "
sql = sql & " VALUES ("
sql = sql & func_ToDate(fch_vuelo) & ","
sql = sql & "" & cod_cia_aerea & ","
sql = sql & "" & cod_vuelo & ","
sql = sql & "'" & cod_depn & "',"
sql = sql & "'" & cod_sdep & "',"
sql = sql & "" & hora_vuelo & ","
sql = sql & "'N',"
sql = sql & "" & cantN & ","
sql = sql & "'" & Tipo & "',"
sql = sql & "'" & locales & "',0)"



Insert_Pax_Volados = Aplicacion.EjecutarDAO(sql)

Else
    Insert_Pax_Volados = False
End If
End Function
Public Function Insert_Destino() As Boolean
Dim sql As String

sql = "INSERT INTO "
sql = sql & " Estadis.Destinos ("
sql = sql & " cod_cia_aerea,"
sql = sql & " cod_vuelo,"
sql = sql & " hora_vuelo,"
sql = sql & " tipo, "
sql = sql & " destino ) "
sql = sql & " VALUES ("
sql = sql & "" & cod_cia_aerea & ","
sql = sql & "" & cod_vuelo & ","
sql = sql & "" & hora_vuelo & ","
sql = sql & "'" & Tipo & "',"
sql = sql & "'" & Destino & "')"

Insert_Destino = Aplicacion.EjecutarDAO(sql)

End Function
Public Property Get nacionalidad() As Integer
nacionalidad = i_nacionalidad
End Property
Public Property Let nacionalidad(ByVal vNewValue As Integer)
i_nacionalidad = vNewValue
End Property
Public Property Get Tipo() As String
Tipo = s_Tipo
End Property
Public Property Let Tipo(ByVal vNewValue As String)
s_Tipo = vNewValue
End Property
Public Function update_destino() As Boolean

Dim sql As String

sql = "UPDATE "
sql = sql & " estadis.destinos SET "
sql = sql & " destino = '" & Destino & "' ,"
sql = sql & " hora_vuelo = " & hora_vuelo
sql = sql & " WHERE cod_cia_aerea = " & cod_cia_aerea
sql = sql & " and cod_vuelo = " & cod_vuelo
sql = sql & " and tipo = '" & Tipo & "' "

update_destino = Aplicacion.EjecutarDAO(sql)

End Function

Public Function update_pax_volados() As Boolean

Dim sql As String

sql = "UPDATE "
sql = sql & " estadis.pax_volados SET "
sql = sql & " cod_depn = '" & cod_depn & "',"
sql = sql & " cod_sdep = '" & cod_sdep & "',"
sql = sql & " hora_vuelo = " & hora_vuelo & ","
sql = sql & " cantidad = " & (cantE + cantT) & ","
sql = sql & " tipo = '" & Tipo & "',"
sql = sql & " local = '" & locales & "',"
sql = sql & " transito = " & cantT
sql = sql & " WHERE fch_vuelo = " & func_ToDate(fch_vuelo)
sql = sql & " and cod_cia_aerea = " & cod_cia_aerea
sql = sql & " and cod_vuelo = " & cod_vuelo
sql = sql & " and nacionalidad = 'E' "

update_pax_volados = Aplicacion.EjecutarDAO(sql)


sql = "UPDATE "
sql = sql & " estadis.pax_volados SET "
sql = sql & " cod_depn = '" & cod_depn & "',"
sql = sql & " cod_sdep = '" & cod_sdep & "',"
sql = sql & " hora_vuelo = " & hora_vuelo & ","
sql = sql & " cantidad = " & cantN & ","
sql = sql & " tipo = '" & Tipo & "',"
sql = sql & " local = '" & locales & "',"
sql = sql & " transito = " & cantT
sql = sql & " WHERE fch_vuelo = " & func_ToDate(fch_vuelo)
sql = sql & " and cod_cia_aerea = " & cod_cia_aerea
sql = sql & " and cod_vuelo = " & cod_vuelo
sql = sql & " and nacionalidad = 'N' "

update_pax_volados = Aplicacion.EjecutarDAO(sql)

End Function
Public Function Delete_Destino() As Boolean

Dim sql As String

sql = "DELETE "
sql = sql & " FROM ESTADIS.destinos  "
sql = sql & " WHERE cod_cia_aerea = " & cod_cia_aerea
sql = sql & " and cod_vuelo = " & cod_vuelo
sql = sql & " and tipo = '" & Tipo & "' "

Delete_Destino = Aplicacion.EjecutarDAO(sql)

End Function

Public Function delete_pax_volados() As Boolean

Dim sql As String

sql = "DELETE "
sql = sql & " FROM ESTADIS.pax_volados  "
sql = sql & " WHERE fch_vuelo = " & func_ToDate(fch_vuelo)
sql = sql & " and cod_cia_aerea = " & cod_cia_aerea
sql = sql & " and cod_vuelo = " & cod_vuelo
sql = sql & " and nacionalidad = 'E' "

delete_pax_volados = Aplicacion.EjecutarDAO(sql)


sql = "DELETE "
sql = sql & " FROM estadis.pax_volados  "
sql = sql & " WHERE fch_vuelo = " & func_ToDate(fch_vuelo)
sql = sql & " and cod_cia_aerea = " & cod_cia_aerea
sql = sql & " and cod_vuelo = " & cod_vuelo
sql = sql & " and nacionalidad = 'N' "

delete_pax_volados = Aplicacion.EjecutarDAO(sql)

End Function

Public Property Get cantE() As Integer
cantE = i_cantE
End Property

Public Property Get cantT() As Integer
cantT = i_cantT
End Property
Public Property Let cantE(ByVal vNewValue As Integer)
i_cantE = vNewValue
End Property

Public Property Let cantT(ByVal vNewValue As Integer)
i_cantT = vNewValue
End Property
Public Property Get cantN() As Integer
cantN = i_cantN
End Property

Public Property Let cantN(ByVal vNewValue As Integer)
i_cantN = vNewValue
End Property


Public Property Get locales() As Variant
locales = s_Local
End Property

Public Property Let locales(ByVal vNewValue As Variant)
    s_Local = vNewValue
End Property
