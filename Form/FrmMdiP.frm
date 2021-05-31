VERSION 5.00
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "InterBaires S.A."
   ClientHeight    =   6510
   ClientLeft      =   330
   ClientTop       =   1080
   ClientWidth     =   7650
   Icon            =   "FrmMdiP.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu MnuGral 
      Caption         =   "&Estadistico"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuEst 
         Caption         =   "Ventas por &Local Grupo"
         Index           =   0
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Ventas por &Grupo Local"
         Index           =   1
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Ventas por &Cajero"
         Index           =   2
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Ventas por &Totales"
         Index           =   3
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Ventas por Prov-&Marc-Prod"
         Index           =   5
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Seguimiento de &Productos"
         Index           =   6
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Ventas por Com&pañía Aérea"
         Index           =   7
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Venta &Horaria"
         Index           =   8
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Resumen Acumulado Diario"
         Index           =   9
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Facturación por &Nacionalidad"
         Index           =   10
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Venta por Hora de PROCESO"
         Index           =   11
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Participación por monedas"
         Index           =   12
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Estimado Unidades vs. Ventas"
         Index           =   13
         Begin VB.Menu mnuvtavsetim 
            Caption         =   "Monitoreo "
            Index           =   1
         End
         Begin VB.Menu mnuvtavsetim 
            Caption         =   "General"
            Index           =   2
         End
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Estadisticas de D.F.Point"
         Index           =   14
         Begin VB.Menu mnuDFP 
            Caption         =   "Indicador General"
            Index           =   1
         End
         Begin VB.Menu mnuDFP 
            Caption         =   "por Nacionalidad"
            Index           =   2
         End
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Tickets de Ventas"
         Index           =   15
      End
      Begin VB.Menu mnuEst 
         Caption         =   "Seguimiento de Lanzamientos"
         Index           =   16
      End
   End
   Begin VB.Menu MnuGral 
      Caption         =   "&Consultas Comparativas"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnuvs 
         Caption         =   "por Diario-&Rubro-Comitente"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuvs 
         Caption         =   "por &Local-Rubro-Comitente"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuvs 
         Caption         =   "por &Diario"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuvs 
         Caption         =   "&Indicadores"
         Index           =   3
      End
      Begin VB.Menu mnuvs 
         Caption         =   "Indicadores Llegadas / Salidas"
         Index           =   4
      End
      Begin VB.Menu mnuvs 
         Caption         =   "Indicadores por Locales"
         Index           =   5
      End
      Begin VB.Menu mnuvs 
         Caption         =   "Comparativo Ventas &Anteriores"
         Index           =   6
      End
      Begin VB.Menu mnuvs 
         Caption         =   "Resumen Dia a Dia"
         Index           =   7
      End
      Begin VB.Menu mnuvs 
         Caption         =   "Pasajeros &Viajados por Nacionalidad"
         Index           =   8
      End
      Begin VB.Menu mnuvs 
         Caption         =   "Ventas por Rubro-&Local"
         Index           =   9
      End
      Begin VB.Menu mnuvs 
         Caption         =   "Ventas por Local Ru&bro"
         Index           =   10
      End
      Begin VB.Menu mnuvs 
         Caption         =   "Ventas por Rubro - Nacionalidad "
         Index           =   11
      End
      Begin VB.Menu mnuvs 
         Caption         =   "Ventas por Nacionalidad - Rubro"
         Index           =   12
      End
      Begin VB.Menu mnuvs 
         Caption         =   "Ventas por Rubro - Proveedor"
         Index           =   13
      End
      Begin VB.Menu mnuvs 
         Caption         =   "Venta por metro Cuadrado"
         Index           =   14
      End
      Begin VB.Menu mnuvs 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuvs 
         Caption         =   "&Modelos"
         Index           =   16
         Begin VB.Menu mnuEstimado 
            Caption         =   "&Inicial"
            Index           =   0
            Begin VB.Menu mnuInic 
               Caption         =   "&Insertar Nuevos"
               Index           =   0
            End
            Begin VB.Menu mnuInic 
               Caption         =   "&Modificación y Consultas"
               Index           =   1
            End
         End
         Begin VB.Menu mnuEstimado 
            Caption         =   "&Nuevos Modelos"
            Index           =   2
         End
         Begin VB.Menu mnuEstimado 
            Caption         =   "&Modelos Existentes"
            Index           =   3
         End
      End
   End
   Begin VB.Menu MnuGral 
      Caption         =   "Ventas &InFlight"
      Index           =   2
   End
   Begin VB.Menu MnuGral 
      Caption         =   "Ventas &Online"
      Index           =   3
      Begin VB.Menu mnuOnLine 
         Caption         =   "Datos Generales"
         Index           =   1
      End
      Begin VB.Menu mnuOnLine 
         Caption         =   "Control de Tickets"
         Index           =   2
      End
      Begin VB.Menu mnuOnLine 
         Caption         =   "Control Horario"
         Index           =   3
      End
   End
   Begin VB.Menu MnuGral 
      Caption         =   "&Captación"
      Index           =   5
      Visible         =   0   'False
      Begin VB.Menu mnuCap 
         Caption         =   "&Viajados"
         Index           =   0
         Begin VB.Menu mnuviajados 
            Caption         =   "&Insertar"
            Index           =   0
         End
         Begin VB.Menu mnuviajados 
            Caption         =   "&Modificaciones y Consultas"
            Index           =   1
         End
         Begin VB.Menu mnuviajados 
            Caption         =   "Informar &FIN"
            Index           =   2
         End
         Begin VB.Menu mnuviajados 
            Caption         =   "Destinos"
            Index           =   3
         End
      End
      Begin VB.Menu mnuCap 
         Caption         =   "&Generación de Comparativo"
         Index           =   1
      End
      Begin VB.Menu mnuCap 
         Caption         =   "&Comparativo Atendidos/Volados"
         Index           =   2
      End
      Begin VB.Menu mnuCap 
         Caption         =   "&Listados"
         Index           =   3
      End
   End
   Begin VB.Menu MnuGral 
      Caption         =   "&Premios"
      Index           =   6
      Visible         =   0   'False
      Begin VB.Menu mnuCom 
         Caption         =   "Conc&ursos"
         Index           =   4
         Begin VB.Menu mnuConc 
            Caption         =   "Adm. de &Concursos-Proveedores"
            Index           =   0
         End
         Begin VB.Menu mnuConc 
            Caption         =   "Asignación de &Productos"
            Index           =   1
         End
         Begin VB.Menu mnuConc 
            Caption         =   "Adm de Productividad"
            Index           =   2
         End
         Begin VB.Menu mnuConc 
            Caption         =   "Cálculo de Productividad"
            Index           =   3
         End
         Begin VB.Menu mnuConc 
            Caption         =   "-"
            Index           =   8
         End
      End
      Begin VB.Menu mnuCom 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuCom 
         Caption         =   "&Informe Comision Variable"
         Index           =   10
      End
      Begin VB.Menu mnuCom 
         Caption         =   "Informe Concursos"
         Index           =   11
      End
   End
   Begin VB.Menu MnuGral 
      Caption         =   "&Admistración"
      Index           =   7
      Visible         =   0   'False
      Begin VB.Menu mnuAdm 
         Caption         =   "&Equipos"
         Index           =   0
         Begin VB.Menu mnuEquipo 
            Caption         =   "Equipos &Nuevos "
            Index           =   0
         End
         Begin VB.Menu mnuEquipo 
            Caption         =   "Equipos &Existentes"
            Index           =   1
         End
      End
      Begin VB.Menu mnuAdm 
         Caption         =   "&Personal"
         Index           =   1
         Begin VB.Menu mnuPersona 
            Caption         =   "&Asignar Nuevo Personal"
            Index           =   0
         End
         Begin VB.Menu mnuPersona 
            Caption         =   "&Consultar Personal"
            Index           =   1
         End
      End
      Begin VB.Menu mnuAdm 
         Caption         =   "&Cajeros"
         Index           =   2
         Begin VB.Menu MnuCajeros 
            Caption         =   "Cajeros &Nuevos"
            Index           =   0
         End
         Begin VB.Menu MnuCajeros 
            Caption         =   "Cajeros &Existentes"
            Index           =   1
         End
      End
      Begin VB.Menu mnuAdm 
         Caption         =   "&Ventas"
         Index           =   3
         Begin VB.Menu mnuVta 
            Caption         =   "Act. Ventas del Interior"
            Index           =   0
         End
      End
      Begin VB.Menu mnuAdm 
         Caption         =   "Ausencias"
         Index           =   4
         Begin VB.Menu mnuAus 
            Caption         =   "Carga Ausencias"
            Index           =   0
         End
      End
      Begin VB.Menu mnuAdm 
         Caption         =   "Interface SAP"
         Index           =   5
      End
      Begin VB.Menu mnuAdm 
         Caption         =   "Lanzamientos"
         Index           =   6
      End
      Begin VB.Menu mnuAdm 
         Caption         =   "Linea/Proveedores"
         Index           =   7
      End
      Begin VB.Menu mnuAdm 
         Caption         =   "Usuarios"
         Index           =   8
      End
      Begin VB.Menu mnuAdm 
         Caption         =   "ABM Sectores"
         Index           =   9
      End
   End
   Begin VB.Menu MnuGral 
      Caption         =   "Ve&ntana"
      Index           =   10
      WindowList      =   -1  'True
      Begin VB.Menu mnuTana 
         Caption         =   "&Estado Pax"
         Index           =   0
      End
      Begin VB.Menu mnuTana 
         Caption         =   "&Estado Vtas Interior"
         Index           =   1
      End
      Begin VB.Menu mnuTana 
         Caption         =   "&Criterios"
         Index           =   2
      End
   End
   Begin VB.Menu MnuGral 
      Caption         =   "&Salir"
      Index           =   11
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub L_NuevaConexion()
Dim i
Dim Frms() As String

On Error GoTo ErrCon:

If MsgBox("Quiere iniciar una nueva conexión ", vbQuestion + vbOKCancel, "ATENCION") = vbOK Then
ReDim Frms(frmPrincipal.lstForms.ListCount)
For i = 0 To frmPrincipal.lstForms.ListCount - 1
    Frms(i + 1) = frmPrincipal.lstForms.List(i)
Next

For i = 1 To UBound(Frms)
    Select Case Frms(i)
    Case "frmLGC"
        Unload frmLGC
    Case "frmGLC"
        Unload frmGLC
    Case "frmCaj"
        Unload frmCaj
    Case "frmTot"
        Unload frmTot
    Case "frmComp"
        Unload frmComp
    Case "frmPMPG"
        Unload frmPMPG
    Case "frmProd"
        Unload frmProd
    Case "frmCiaAerea"
        Unload frmCiaAerea
    Case "frmVsRub"
        Unload frmVsRub
    Case "frmIndic"
        Unload frmIndic
    Case "frmVsDia"
        Unload frmVsDia
    Case "frmPerAnual"
        Unload frmPerAnual
    Case "frmNacion"
        Unload frmNacion
        
  End Select
    
Next

Call Main
End If
ErrCon:
    Exit Sub
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'Unload frmTM
End
End Sub


Private Sub mnuAdm_Click(Index As Integer)
Select Case Index
    Case 5
    frmInterfaceSAP.Show
    Case 6
    FrmAdmLanzamientos.Show
    Case 7
    FrmAdmMonitoreoLinea.Show
    Case 8
    FrmAdmUsuario.Show
    Case 9
    frmAdmSectorProveedor.Show

End Select
End Sub

Private Sub mnuAus_Click(Index As Integer)
    FrmAdmAusen.Show
End Sub

Private Sub MnuCajeros_Click(Index As Integer)

'Select Case Index
'   Case 0
'     FrmAdmCajero.Altas
'   Case 1
'     FrmAdmCajero.Modificacion
'End Select

FrmAdmCajero.Show

End Sub

Private Sub mnuCap_Click(Index As Integer)

Select Case Index
  Case 1
    FrmComparativo.Show 1
  Case 2
    FrmVersus.Show 1
  Case 3
    frmInformes.Show 1
    
End Select

End Sub

Private Sub mnuCom_Click(Index As Integer)

Select Case Index
    Case 10
        frmInsentivo.Show
    Case 11
        frmConsultaConcurso.Show
        'MsgBox "Consulta suspendida por modificaciones ", vbInformation + vbOKOnly, "ATENCION"
        
End Select

End Sub
Private Sub mnuConc_Click(Index As Integer)

Select Case Index
    Case 0
        FrmAdmPxC.Show
    Case 1
        FrmAdmConcurso.Show
    Case 2
        FrmAdmProductividad.Show
    Case 3
        frmExeProductividad.Show

End Select

End Sub

Private Sub mnuDFP_Click(Index As Integer)
Select Case Index
    Case 1
         frmPuntosGral.Show
    Case 2
         frmPuntosNac.Show
End Select
End Sub

Private Sub mnuEquipo_Click(Index As Integer)

Select Case Index
    Case 0
        FrmAdmEquip.altas
    Case 1
        FrmAdmEquip.modificacion
End Select
End Sub

Private Sub mnuEst_Click(Index As Integer)
Select Case Index
    Case 0
        frmLGC.Show
    Case 1
        frmGLC.Show
    Case 2
        frmCaj.Show
    Case 3
        frmTot.Show
    Case 5
        frmPMPG.Show
    Case 6
        frmProd.Show
    Case 7
        frmCiaAerea.Show
    Case 8
        frmVtaHora.ModoConsulta 1
    Case 9
        frmAcum.Show
    Case 10
        frmLocNac.Show
    Case 11
        frmVtaHora.ModoConsulta 2
    Case 12
        frmMonedas.Show
    Case 15
        FrmConsultaTicket.Show
    Case 16
        frmVentasLanzamientos.Show
End Select

End Sub

Private Sub mnuEstimado_Click(Index As Integer)
Select Case Index
    Case 2
        FrmEstimN.altas
    Case 3
        FrmEstimN.modificacion
End Select
End Sub

Private Sub mnuGral_Click(Index As Integer)
Select Case Index
    Case 8
        L_NuevaConexion
    Case 2
        frmResumenDiarioIFL.Show
    Case 3
        'FrmOnline.Show
    Case 11
    End
End Select
End Sub

Private Sub mnuInic_Click(Index As Integer)
Select Case Index
    Case 0
        FrmAdmModEsp.altas
    Case 1
        FrmAdmModEsp.modificacion
End Select

End Sub

Private Sub mnuOnLine_Click(Index As Integer)
Select Case Index
Case 1
     FrmOnline.Show
Case 2
     frmControlTicket.Show
Case 3
     frmOnlineHora.Show
End Select
End Sub

Private Sub mnuPersona_Click(Index As Integer)
Select Case Index
    Case 0
        'FrmAdmPersonas.Altas
        FrmAdmEmpleado.altas
    Case 1
        'FrmAdmPersonas.Modificacion
        FrmAdmEmpleado.modificacion
End Select
End Sub

Private Sub mnuTana_Click(Index As Integer)

Select Case Index
    Case 0
        frmAbaut.Muestra
    Case 1
        frmAbautTicket.Muestra
    Case 2
        frmCrit.Show 1
End Select

End Sub

Private Sub mnuviajados_Click(Index As Integer)

Select Case Index
    Case 0
      FrmAdmDeFlight.altas
    Case 1
      FrmAdmDeFlight.modificacion
    Case 2
      frmAbaut.Seteo
    Case 3
      FrmAdmDestino.Show
End Select

End Sub


Private Sub mnuVs_Click(Index As Integer)
Select Case Index
     Case 0
        frmVsRub.Show
    Case 1
        frmVsLocRub.Show
    Case 2
        frmVsDia.Show
    Case 3
        frmIndic_t.Show
    Case 4
        frmIndicLS.Show
    Case 5
        frmIndicLoc.Show
    Case 6
        frmComp.Show
    Case 7
        frmResumenDiario.Show
    Case 8
        frmNacion_ES.Show
    Case 9
        frmVtaRubroLocal.Show
    Case 10
        frmVtaLocalRubro.Show
    Case 11
        frmVtaRubroNacion.Show
    Case 12
        frmVtaNacionRubro.Show
        'MsgBox "Disculpe : pantalla momentaneamente deshabilitada por actualización", vbOKOnly + vbExclamation, "Atención"
    Case 13
        frmVtaRubroProv.Show
    Case 14
        frmVentaMetros.Show
        
    
End Select
End Sub


Private Sub mnuVta_Click(Index As Integer)
    Select Case Index
        Case 0
            frmAbautTicket.Seteo
    End Select
    
End Sub


Private Sub mnuvtavsetim_Click(Index As Integer)

Select Case Index
    Case 1
        frmMonitoreoLinea.Show
    Case 2
        frmMonitoreo.Show
End Select
End Sub


