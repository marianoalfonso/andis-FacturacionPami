Attribute VB_Name = "Inicio"
'DECLARACIONES DE VARIABLES GLOBALES
Global Gstr_Cuit As String
Global Gstr_RazonSocial As String
Global Gstr_TipoPrestador As String
Global Gstr_Provincia As String

Global Gstr_PeriodoFacturado As String
Global Gstr_NumeroFactura As String

Global Gstr_FacturacionCerrada As Boolean
Global exp_EstadoFacturacion As String

Global Gstr_NombreArchivoFacturacionAbierta As String

Global Gstr_DestinationPath As String

Global Const cFormatoFecha As String = "dd/mm/yyyy"
Global Const cFormatoNumero As String = "###,###"      ' "###"
Global Const cFormatoMoneda As String = "###,###.##"
' La cantidad de cifras a tener en cuenta en los números
Global Const cCuantasCifras As Long = 20&


Sub main()
    Setear_Configuracion
    frmIndex.Show
End Sub

