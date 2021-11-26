VERSION 5.00
Begin VB.MDIForm frmIndex 
   BackColor       =   &H8000000C&
   Caption         =   "SISTEMA DE CARGA DE FACTURACIONES - Programa Federal Incluir Salud"
   ClientHeight    =   8790
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12510
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnAjustes 
      Caption         =   "Ajustes"
      Index           =   2
      Begin VB.Menu mnConfiguracionInicial 
         Caption         =   "Configuración Inicial"
         Index           =   23
      End
   End
   Begin VB.Menu mnFacturacion 
      Caption         =   "Facturación"
      Index           =   3
      Begin VB.Menu mnNuevaFacturacion 
         Caption         =   "Generar/Cerrar Facturación"
         Index           =   32
      End
      Begin VB.Menu mnFacturacionesHistoricas 
         Caption         =   "Reimprimir/Borrar Facturación"
         Index           =   33
      End
   End
   Begin VB.Menu mnAbout 
      Caption         =   "Acerca de ..."
   End
   Begin VB.Menu mnSalir 
      Caption         =   "Salir"
      Index           =   4
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mensaje As String

Private Sub MDIForm_Load()
    Call Chequeo_ConfiguracionInicial
End Sub

Private Sub mnAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnConfiguracionInicial_Click(Index As Integer)
    frmConfiguracionInicial.Show
End Sub

'CHEQUEAMOS SI ESTA COMPLETA LA CONFIGURACION INICIAL DEL SISTEMA
Sub Chequeo_ConfiguracionInicial()
    If Gstr_Cuit = "" Or Gstr_RazonSocial = "" Or Gstr_TipoPrestador = "" Or Gstr_Provincia = "" Then
        MsgBox "Debe completar la configuración inicial del sistema.", vbInformation
        frmConfiguracionInicial.Show
    End If
End Sub


Private Sub mnFacturacionesHistoricas_Click(Index As Integer)
    frmListadoFacturaciones.BorderStyle = 4
    frmListadoFacturaciones.cmdSeleccionar.Visible = False
    frmListadoFacturaciones.cmdImprimir.Visible = True
    frmListadoFacturaciones.Show
End Sub

Private Sub mnNuevaFacturacion_Click(Index As Integer)
Dim sNombreArchivo As String
    If Estado_Facturacion() Then
        'HAY UNA FACTURACION ABIERTA EN CURSO
        Gstr_NombreArchivoFacturacionAbierta = Trim(ConfigGet("DATOS_FACTURACION", "File"))
        MsgBox "El archivo de facturación " & Gstr_NombreArchivoFacturacionAbierta & " no esta cerrada. Debe exportarla o eliminarla.", vbInformation
        Gstr_FacturacionCerrada = False
      Else
        Gstr_FacturacionCerrada = True
        Gstr_NombreArchivoFacturacionAbierta = "null"
        Mensaje = "Ingrese el periodo a Facturar:" & Chr(13) & "formato: MM-AAAA" & Chr(13) & "ejemplo: 09-2012"
        Gstr_PeriodoFacturado = InputBox(Mensaje, "Periodo de Facturación")
        If Not Validar_Periodo(Gstr_PeriodoFacturado) Then
            Mensaje = "Periodo no válido." & Chr(13) & Chr(13) & "Verifique el formato/periodo"
            MsgBox Mensaje, vbCritical
            Exit Sub
        End If
        Mensaje = "Ingrese el número de factura: (formato: 0000-00000000)" & Chr(13) & Chr(13) & "ejemplos: 0001-00000658" & Chr(13) & "               1-658"
        Gstr_NumeroFactura = InputBox(Mensaje, "Número de factura")
        
        'VALIDA LA FACTURA
        If Validar_Factura(Gstr_NumeroFactura) Then
            sNombreArchivo = Trim(Generar_Archivo_Exportacion(Trim(Gstr_TipoPrestador), Trim(Gstr_PeriodoFacturado), Trim(Gstr_Cuit), Trim(Gstr_NumeroFactura), Trim(exp_EstadoFacturacion)))
            If ArchivoExiste(Trim(App.Path & "\Facturaciones Cerradas\C" & Trim(sNombreArchivo) & ".txt")) Then
                MsgBox "Ya existe una facturacion para este Período con este número de factura." & Chr(13) & Chr(13) & _
                        "Si desea volver a facturar este periodo con este numero de factura" & Chr(13) & _
                        "deberá eliminarla primero desde la opción Reimprimir/Borrar Facturación.", vbCritical
                Exit Sub
            End If
          Else
            Mensaje = "Factura no válida." & Chr(13) & Chr(13) & "Verifique el formato por favor."
            MsgBox Mensaje, vbCritical
            Exit Sub
        End If
    End If
    
    frmFacturacion.Show
End Sub

Private Sub mnSalir_Click(Index As Integer)
    End
End Sub
