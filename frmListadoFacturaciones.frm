VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListadoFacturaciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LISTADO DE FACTURACIONES CERRADAS"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmFacturacionesCerradas 
      BackColor       =   &H80000002&
      Caption         =   "Seleccione una Plantilla de Facturacion"
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   735
         Left            =   1560
         Picture         =   "frmListadoFacturaciones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "frmListadoFacturaciones.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6960
         Width           =   1095
      End
      Begin VB.FileListBox fileFacturaciones 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6510
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   4935
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   735
         Left            =   3960
         Picture         =   "frmListadoFacturaciones.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Seleccionar"
         Height          =   735
         Left            =   360
         Picture         =   "frmListadoFacturaciones.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6960
         Width           =   1095
      End
   End
   Begin VB.OptionButton optSeleccionado 
      Caption         =   "Option1"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   5640
      Width           =   2775
   End
   Begin VB.OptionButton optCancelado 
      Caption         =   "Option1"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   6120
      Width           =   2775
   End
   Begin MSComctlLib.ListView lvFacturacionListado 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1931
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmListadoFacturaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' DEclaración de la Función Api SendMessage
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Dim sBeneficio As String
Dim sNombre As String
Dim sPrestacion As String
Dim sDescripcion As String
Dim sImporteNumerico As Double
Dim sImporte As String


Private Sub Form_Load()
    Me.Top = 200
    Me.Left = 200
    fileFacturaciones.FileName = ""
    fileFacturaciones.Path = App.Path & "\Facturaciones Cerradas"
    fileFacturaciones.Pattern = "*.txt"
End Sub


Private Sub cmdBorrar_Click()
Dim RutaCompleta As String
    If MsgBox("Se eliminará el archivo " & fileFacturaciones.FileName & " de forma definitiva." & Chr(13) & Chr(13) & _
              "¿ desea continuar ?", vbYesNo + vbQuestion) = vbYes Then
        RutaCompleta = App.Path & "\Facturaciones Cerradas\" & Me.fileFacturaciones.FileName
        Call Eliminar_Archivo(RutaCompleta)
        fileFacturaciones.Refresh
    End If
End Sub

Private Sub cmdImprimir_Click()
Dim sLinea As String
Dim sRenglon As String
Dim sArchivo As String
Dim auxTotalPrestacion As String
Dim auxTotalTransporte As String
Dim auxTotalApoyo As String
Dim auxTotalMatricula As String
Dim auxTotalHemodialisis As String
Dim auxTotalGeriatria As String
Dim auxTotalGeneral As Double
Dim auxTotalGeneralString As String
Dim auxArchivoImprimir As String

    auxArchivoImprimir = "Periodo: " & Trim(Mid(fileFacturaciones.FileName, 3, 2) & "-" & Trim(Mid(fileFacturaciones.FileName, 5, 4)) & Chr(13) & "Factura: " & Left(Right(fileFacturaciones.FileName, 17), 13) & Chr(13) & Chr(13) & "Prepare la impresora por favor . . .")
    If MsgBox(auxArchivoImprimir, vbOKCancel + vbQuestion) = vbCancel Then
        Exit Sub
    End If

    With lvFacturacionListado
        ' Las pruebas serán en modo "detalle"
        .View = lvwReport
        ' al seleccionar un elemento, seleccionar la línea completa
        .FullRowSelect = True
        ' Mostrar las líneas de la cuadrícula
        .GridLines = True
        ' No permitir la edición automática del texto
        .LabelEdit = lvwManual
        ' Permitir múltiple selección
        .MultiSelect = True
        ' Para que al perder el foco,
        ' se siga viendo el que está seleccionado
        .HideSelection = False
        .ColumnHeaders.Add , , "Beneficio", 1400, lvwColumnLeft
        .Tag = "Beneficio"
        .ColumnHeaders.Add , , "Nombre", 2200, lvwColumnLeft
        .Tag = "Nombre"
        .ColumnHeaders.Add , , "Prestacion", 1000, lvwColumnCenter
        .Tag = "Prestacion"
        .ColumnHeaders.Add , , "Descripcion", 3400, lvwColumnLeft
        .Tag = "Descripcion"
        .ColumnHeaders.Add , , "Importe", 1100, lvwColumnRight
        .Tag = "Importe"
    End With
    sArchivo = fileFacturaciones.FileName
    Gstr_PeriodoFacturado = ConfigGetEXTENDIDO(sArchivo, "HEADER", "PERIODOFACTURADO")
    Gstr_NumeroFactura = ConfigGetEXTENDIDO(sArchivo, "HEADER", "NUMEROFACTURA")
    sRenglon = Chr(13) & Chr(10)
    Open App.Path & "\Facturaciones Cerradas\" & sArchivo For Input As #1
    Line Input #1, sLinea: Line Input #1, sLinea
    Line Input #1, sLinea: Line Input #1, sLinea
    Line Input #1, sLinea: Line Input #1, sLinea
    Line Input #1, sLinea: Line Input #1, sLinea
    While Not EOF(1)
    'EN LA LINEA 8 COMIENZA EL REGISTRO
        Line Input #1, sLinea:
        If sLinea = "[BOTTOM]" Then
            Select Case Gstr_TipoPrestador
                Case "Discapacidad"
                    auxTotalPrestacion = ConfigGetEXTENDIDO(sArchivo, "BOTTOM", "TotalPrestacion")
                    auxTotalTransporte = ConfigGetEXTENDIDO(sArchivo, "BOTTOM", "TotalTransporte")
                    auxTotalApoyo = ConfigGetEXTENDIDO(sArchivo, "BOTTOM", "TotalApoyo")
                    auxTotalMatricula = ConfigGetEXTENDIDO(sArchivo, "BOTTOM", "TotalMatricula")
                    auxTotalGeneral = Val(ConvertirDecimal(auxTotalPrestacion)) + Val(ConvertirDecimal(auxTotalTransporte)) + Val(ConvertirDecimal(auxTotalApoyo)) + Val(ConvertirDecimal(auxTotalMatricula))
                    auxTotalGeneralString = auxTotalGeneral
                    Call Imprimir_ListView(lvFacturacionListado, sArchivo, Gstr_RazonSocial, Gstr_Cuit, Gstr_PeriodoFacturado, Gstr_NumeroFactura, auxTotalGeneralString, Trim(auxTotalPrestacion), Trim(auxTotalTransporte), Trim(auxTotalApoyo), Trim(auxTotalMatricula))
                Case "Hemodialisis"
                    auxTotalPrestacion = ConfigGetEXTENDIDO(sArchivo, "BOTTOM", "TotalPrestacion")
                    auxTotalTransporte = ConfigGetEXTENDIDO(sArchivo, "BOTTOM", "TotalTransporte")
                    auxTotalGeneral = Val(ConvertirDecimal(auxTotalPrestacion)) + Val(ConvertirDecimal(auxTotalTransporte))
                    auxTotalGeneralString = auxTotalGeneral
                    Call Imprimir_ListView(lvFacturacionListado, sArchivo, Gstr_RazonSocial, Gstr_Cuit, Gstr_PeriodoFacturado, Gstr_NumeroFactura, auxTotalGeneralString, Trim(auxTotalPrestacion), Trim(auxTotalTransporte))
                Case "Geriatria"
                    auxTotalGeneral = Val(ConvertirDecimal(auxTotalPrestacion))
                    auxTotalPrestacion = ConfigGetEXTENDIDO(sArchivo, "BOTTOM", "TotalPrestacion")
                    auxTotalGeneralString = auxTotalGeneral
                    Call Imprimir_ListView(lvFacturacionListado, sArchivo, Gstr_RazonSocial, Gstr_Cuit, Gstr_PeriodoFacturado, Gstr_NumeroFactura, auxTotalGeneralString, Trim(auxTotalPrestacion))
            End Select
            Exit Sub
          Else
            sBeneficio = sLinea
        End If
        Line Input #1, sLinea: sNombre = sLinea
        Line Input #1, sLinea: sPrestacion = sLinea
        Line Input #1, sLinea: sDescripcion = sLinea
        Line Input #1, sLinea: sImporte = sLinea
        Call Agregar_Item(sBeneficio, sNombre, sPrestacion, sDescripcion, sImporte)
        Line Input #1, sLinea
        If sLinea <> "endline" Then
            MsgBox "Error imprimiendo el archivo.", vbInformation
        End If
    Wend
    Close #1
End Sub

Private Sub Agregar_Item(Beneficio As String, Nombre As String, TipoPrestacion As String, Descripcion As String, Importe As String)
    Dim i As Long
    With lvFacturacionListado.ListItems.Add(, , Beneficio)
        .SubItems(1) = Nombre
        .SubItems(2) = TipoPrestacion
        .SubItems(3) = Descripcion
        .SubItems(4) = Importe
    End With
    Call Ordenar_Items
'    Call Configurar_FrameCarga
End Sub

Private Sub cmdCancelar_Click()
    optCancelado.Value = True
    optSeleccionado.Value = False
    Me.Hide
End Sub

Private Sub cmdSeleccionar_Click()
    optCancelado.Value = False
    optSeleccionado.Value = True
    Me.Hide
End Sub

Sub Ordenar_Items()
    Call SendMessage(Me.hwnd, WM_SETREDRAW, 0&, 0&)
    With lvFacturacionListado
        .SortOrder = lvwAscending
        .SortKey = .ColumnHeaders(1).Index
        .Sorted = True
    End With
    Call SendMessage(Me.hwnd, WM_SETREDRAW, 1&, 0&)
    lvFacturacionListado.Refresh
End Sub
