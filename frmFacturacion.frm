VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacturacion 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Facturación"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11835
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H8000000A&
      Caption         =   "Salir"
      Height          =   975
      Left            =   10440
      Picture         =   "frmFacturacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Frame frmTotalesDiscapacidad 
      Caption         =   "TOTALES"
      Height          =   2895
      Left            =   9720
      TabIndex        =   25
      Top             =   240
      Width           =   1935
      Begin VB.TextBox txtTotalMatricula 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtTotalApoyo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtTotalTransporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtTotalPrestacion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Apoyo"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Matricula"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Transporte"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Prestacion"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdCargarPlantilla 
      Caption         =   "Cargar Plantilla"
      Height          =   855
      Left            =   9840
      Picture         =   "frmFacturacion.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdEliminarItem 
      Caption         =   "Eliminar Item"
      Height          =   735
      Left            =   9840
      Picture         =   "frmFacturacion.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdEditarItem 
      Caption         =   "Editar Item"
      Height          =   735
      Left            =   9840
      Picture         =   "frmFacturacion.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdBorrarFacturacionActual 
      Caption         =   "Borrar Facturación en curso"
      Height          =   855
      Left            =   9840
      Picture         =   "frmFacturacion.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar facturacion para su posterior edición"
      Height          =   975
      Left            =   360
      Picture         =   "frmFacturacion.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7320
      Width           =   3375
   End
   Begin VB.CommandButton cmdExportar 
      BackColor       =   &H80000003&
      Caption         =   "Cerrar e Imprimir Facturación"
      Height          =   975
      Left            =   3840
      Picture         =   "frmFacturacion.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7320
      Width           =   3375
   End
   Begin MSComctlLib.ListView lvFacturacion 
      Height          =   3855
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6800
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "CARGA DE CONSUMOS"
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   9375
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         MaxLength       =   8
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgregarPrestacion 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   8160
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   4320
         MaxLength       =   60
         TabIndex        =   9
         Top             =   720
         Width           =   4815
      End
      Begin VB.ComboBox cmbTipoPrestacion 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtBeneficio 
         Height          =   285
         Left            =   1320
         MaxLength       =   14
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblImporte 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   3720
         TabIndex        =   12
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion:"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Prest:"
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblBeneficio 
         AutoSize        =   -1  'True
         Caption         =   "Beneficio:"
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DATOS DE CABECERA"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9375
      Begin VB.Label lblPeriodoFacturado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "periodofacturado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblNumeroFactura 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         TabIndex        =   17
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label lblCuit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "cuit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label lblPrestador 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "prestador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame frmTotalesGeriatria 
      Caption         =   "TOTALES"
      Height          =   2895
      Left            =   9720
      TabIndex        =   40
      Top             =   240
      Width           =   1935
      Begin VB.TextBox txtTotalGeriatria 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total Geriatria"
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Frame frmTotalesHemodialisis 
      Caption         =   "TOTALES"
      Height          =   2895
      Left            =   9720
      TabIndex        =   35
      Top             =   240
      Width           =   1935
      Begin VB.TextBox txtTotalHemodialisis 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtTotalTransporteHemodialisis 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Total Hemodialisis"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total Transporte"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   1170
      End
   End
   Begin VB.Label lblMensaje 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   14
      Top             =   3000
      Width           =   9165
   End
End
Attribute VB_Name = "frmFacturacion"
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
Dim dTotalPrestacion As Double
Dim dTotalTransporte As Double
Dim dTotalApoyo As Double
Dim dTotalMatricula As Double
Dim iConteoCaracteres As Integer
Dim bActualizando As Boolean
Dim bEditando As Boolean
Dim auxTotalImprimir As String
Dim sTotalGeneral As Double
Dim saux_TotalGeneral As String


Private Sub Form_Load()
    Me.Top = 100
    Me.Left = 100
    Call Configurar_Formulario
End Sub

Sub Configurar_Formulario()
    bEditando = False
    Call Configurar_Cabecera
    Call Configurar_FrameCarga
    Call Configurar_lvFacturacion
    Call Mostrar_Mensaje("El listado se ordena alfabeticamente por nombre para su mejor organización")
    Call Habilitar_FrameTotales
    Call Inicializar_Totales
    Call Definir_ToolTipText
    If Not Gstr_FacturacionCerrada Then
        Call Cargar_Facturacion_Abierta
    End If
End Sub

Sub Configurar_Cabecera()
    lblPrestador.Caption = Gstr_RazonSocial
    lblCuit.Caption = Gstr_Cuit
    lblPeriodoFacturado.Caption = Gstr_PeriodoFacturado
    lblNumeroFactura.Caption = Gstr_NumeroFactura
End Sub

Sub Configurar_FrameCarga()
    txtBeneficio.Text = "": txtNombre.Text = "": txtDescripcion.Text = "": txtImporte = ""
    Call Fill_ComboPrestacion
End Sub

'MUESTRA EL FRAME DE TOTALES CORRESPONDIENTE
Sub Habilitar_FrameTotales()
    Select Case Gstr_TipoPrestador
        Case "Discapacidad"
            frmTotalesDiscapacidad.Visible = True
            frmTotalesHemodialisis.Visible = False
            frmTotalesGeriatria.Visible = False
        Case "Hemodialisis"
            frmTotalesDiscapacidad.Visible = False
            frmTotalesHemodialisis.Visible = True
            frmTotalesGeriatria.Visible = False
        Case "Geriatria"
            frmTotalesDiscapacidad.Visible = False
            frmTotalesHemodialisis.Visible = False
            frmTotalesGeriatria.Visible = True
    End Select
End Sub

'SEGUN EL TIPO DE PRESTADOR LLENA EL COMBO DE OPCIONES
Sub Fill_ComboPrestacion()
    Select Case Gstr_TipoPrestador
        Case "Discapacidad"
            cmbTipoPrestacion.Clear
            cmbTipoPrestacion.AddItem "Prestacion", 0
            cmbTipoPrestacion.AddItem "Transporte", 1
            cmbTipoPrestacion.AddItem "Apoyo", 2
            cmbTipoPrestacion.AddItem "Matricula", 3
            cmbTipoPrestacion.ListIndex = -1
        Case "Hemodialisis"
            cmbTipoPrestacion.Clear
            cmbTipoPrestacion.AddItem "Hemodialisis", 0
            cmbTipoPrestacion.AddItem "Transporte", 1
            cmbTipoPrestacion.ListIndex = -1
        Case "Geriatria"
            cmbTipoPrestacion.Clear
            cmbTipoPrestacion.AddItem "Geriatria", 0
            cmbTipoPrestacion.ListIndex = 0
    End Select
End Sub

Sub Configurar_lvFacturacion()
    'CONFIGURACION DEL LISTVIEW PRINCIPAL
    With lvFacturacion
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

End Sub

Sub Inicializar_Totales()
    Select Case Gstr_TipoPrestador
        Case "Discapacidad"
            txtTotalPrestacion.Text = "0,00"
            txtTotalTransporte.Text = "0,00"
            txtTotalApoyo.Text = "0,00"
            txtTotalMatricula.Text = "0,00"
        Case "Hemodialisis"
            txtTotalHemodialisis.Text = "0,00"
            txtTotalTransporteHemodialisis.Text = "0,00"
        Case "Geriatria"
            txtTotalGeriatria.Text = "0,00"
    End Select
End Sub

Sub Definir_ToolTipText()
    cmdCargarPlantilla.ToolTipText = "Permite usar como plantilla una facturacion ya cerrada"
    cmdExportar.ToolTipText = "Exporta y cierra la facturación actual, luego de este proceso no se permite editarla"
    cmdGuardar.ToolTipText = "Graba la facturación actual para una posterior edición de la misma"
    cmdEditarItem.ToolTipText = "Edita el ítem seleccionado en la grilla de facturación"
    cmdEliminarItem.ToolTipText = "Elimina el ítem seleccionado de la grilla de facturación"
    cmdBorrarFacturacionActual.ToolTipText = "Elimina de forma irreversible el archivo actual de facturación sin cerrar"
    
    
End Sub

Private Sub cmdAgregarPrestacion_Click()
'CHEQUEAR DUPLICIDAD DE REGISTRO (BENEFICIO + TIPO PRESTACION)
    If bEditando Then
        Call Habilitar_Formulario(True)
    End If
    If Validar_Datos Then
        sBeneficio = Trim(txtBeneficio.Text)
        sNombre = Trim(txtNombre.Text)
        sPrestacion = Trim(cmbTipoPrestacion.Text)
        sDescripcion = Trim(txtDescripcion.Text)
        sImporte = Trim(txtImporte.Text)
        Call Agregar_Item(sBeneficio, sNombre, sPrestacion, sDescripcion, sImporte)
        Call Actualizar_Totales(False)
    End If
End Sub


Private Sub Agregar_Item(Beneficio As String, Nombre As String, TipoPrestacion As String, Descripcion As String, Importe As String)
    Dim i As Long
    With lvFacturacion.ListItems.Add(, , Beneficio)
        .SubItems(1) = Nombre
        .SubItems(2) = TipoPrestacion
        .SubItems(3) = Descripcion
        .SubItems(4) = Importe
    End With
    Call Ordenar_Items
    Call Configurar_FrameCarga
End Sub

Private Sub Reemplazar_Item(Indice As Integer, Beneficio As String, Nombre As String, TipoPrestacion As String, Descripcion As String, Importe As String)
    Dim i As Long
    lvFacturacion.ListItems.Remove (Indice)
    With lvFacturacion.ListItems().Add(, , Beneficio)
        .SubItems(1) = Nombre
        .SubItems(2) = TipoPrestacion
        .SubItems(3) = Descripcion
        .SubItems(4) = Importe
    End With
    Call Ordenar_Items
    Call Configurar_FrameCarga
End Sub


Sub Ordenar_Items()
    Call SendMessage(Me.hwnd, WM_SETREDRAW, 0&, 0&)
    With lvFacturacion
        .SortOrder = lvwAscending
        .SortKey = .ColumnHeaders(1).Index
        .Sorted = True
    End With
    Call SendMessage(Me.hwnd, WM_SETREDRAW, 1&, 0&)
    lvFacturacion.Refresh
End Sub

'MUESTRA UN MENSAJE PASADO COMO CADENA (restringir longitud)
Sub Mostrar_Mensaje(cadena As String)
    lblMensaje.ForeColor = vbRed
    lblMensaje.Caption = Trim(cadena)
End Sub

Private Sub txtBeneficio_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13
        KeyAscii = 0        ' Para que no "pite"
'        SendKeys "{tab}"    ' Envía una pulsación TAB
    Case 8, 13, 48 To 57
    Case Else
        ' No es una tecla numérica, no admitirla
        KeyAscii = 0
    End Select
End Sub


'VALIDA EL TEXTBOX DE IMPORTE
Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    numeros = "0123456789."
    tecla = Chr(KeyAscii)
    If tecla = vbTab Or tecla = vbBack Then Exit Sub
    posi = InStr(1, numeros, tecla)
    If InStr(1, numeros, tecla) <> 0 Then
        If tecla = "." Then
            For cuenta = 1 To Len(txtImporte)
                If Mid(txtImporte, cuenta, 1) = "." Then
                    KeyAscii = 0
                    Exit For
                End If
            Next cuenta
          Else
            posi2 = InStr(1, txtImporte, ".")
            If posi2 > 0 Then
                If posi2 <= Len(txtImporte) - 2 Then
                    KeyAscii = 0
                End If
            End If
        End If
      Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtImporte_LostFocus()
    txtImporte = Format(Val(txtImporte), "#0.00")
End Sub


'VALIDA LOS DATOS DEL FRAME DE CARGA
Function Validar_Datos() As Boolean
Dim Beneficio_Valido As Boolean
Dim Nombre_Valido As Boolean
Dim TipoPrestacion_Valida As Boolean
Dim Descripcion_Valida As Boolean
Dim Importe_Valido As Boolean
    Validar_Datos = True
    Beneficio_Valido = Validar_Beneficio(Trim(Me.txtBeneficio.Text))
    Nombre_Valido = Validar_TextBoxNulo(Trim(Me.txtNombre.Text))
    TipoPrestacion_Valida = Validar_TextBoxNulo(Trim(Me.cmbTipoPrestacion.Text))
'    Descripcion_Valida = Validar_TextBoxNulo(Trim(Me.txtDescripcion.Text))
    Importe_Valido = Validar_TextBoxNulo(Trim(Me.txtImporte.Text))
    If Importe_Valido Then
        Me.txtImporte.Text = Trim(Me.txtImporte.Text)
    End If
    'analizamos cada evaluacion
    If Not Beneficio_Valido Then
        Me.txtBeneficio.BackColor = vbRed
        Validar_Datos = False
      Else
        Me.txtBeneficio.BackColor = vbWhite
    End If
    If Not Nombre_Valido Then
        Me.txtNombre.BackColor = vbRed
        Validar_Datos = False
      Else
        Me.txtNombre.BackColor = vbWhite
    End If
    If Not TipoPrestacion_Valida Then
        Me.cmbTipoPrestacion.BackColor = vbRed
        Validar_Datos = False
      Else
        Me.cmbTipoPrestacion.BackColor = vbWhite
    End If
'    If Not Descripcion_Valida Then
'        Me.txtDescripcion.BackColor = vbRed
'        Validar_Datos = False
'      Else
'        Me.txtDescripcion.BackColor = vbWhite
'    End If
    If Not Importe_Valido Then
        Me.txtImporte.BackColor = vbRed
        Validar_Datos = False
      Else
        Me.txtImporte.BackColor = vbWhite
    End If
End Function


Private Sub cmdGuardar_Click()
Dim sNombreArchivoExportacion As String
Dim iElementos As Integer
Dim iLoop As Integer
Dim iControlExportacion As Boolean
'VARIABLES PARA COMPLETAR LA LINEA DE EXPORTACION
Dim exp_RazonSocial As String
Dim exp_Cuit As String
Dim exp_PeriodoFacturado As String
Dim exp_NumeroFactura As String
Dim exp_Beneficio As String
Dim exp_Nombre As String
Dim exp_TipoPrestacion As String
Dim exp_Detalle As String
Dim exp_Importe As String
    exp_EstadoFacturacion = "A" 'SETEAMOS EL ESTADO DE LA FACTURACION SOBRE LA VARIABLE GLOBAL
    'SI ES UNA ACTUALIZACION, ELIMINAMOS EL ARCHIVO ORIGINAL Y GENERAMOS EL NUEVO CON EL MISMO NOMBRE
    sNombreArchivoExportacion = Generar_Archivo_Exportacion(Trim(Gstr_TipoPrestador), Trim(Gstr_PeriodoFacturado), Trim(Gstr_Cuit), Trim(Gstr_NumeroFactura), Trim(exp_EstadoFacturacion))
    If Actualizando Then
        Kill App.Path & "\" & sNombreArchivoExportacion & ".txt"
    End If
    iElementos = lvFacturacion.ListItems.Count
    'GRABAMOS CABECERA
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, Trim("[STATUS]"), True)
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, Trim("SF=open"), False)
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, Trim("[HEADER]"), False)
    exp_RazonSocial = "PRESTADOR=" & Gstr_RazonSocial
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_RazonSocial, False)
    exp_Cuit = "CUIT=" & Gstr_Cuit
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_Cuit, False)
    exp_PeriodoFacturado = "PERIODOFACTURADO=" & Gstr_PeriodoFacturado
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_PeriodoFacturado, False)
    exp_NumeroFactura = "NUMEROFACTURA=" & Gstr_NumeroFactura
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_NumeroFactura, False)
    iConteoCaracteres = iConteoCaracteres + Len(Trim("[DETAIL]"))
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, "[DETAIL]", False)
    If iControlExportacion Then
        If iElementos > 0 Then
            For iLoop = 1 To iElementos
                exp_Beneficio = Trim(lvFacturacion.ListItems(iLoop))
                exp_Nombre = Trim(lvFacturacion.ListItems(iLoop).SubItems(1))
                exp_TipoPrestacion = Trim(lvFacturacion.ListItems(iLoop).SubItems(2))
                exp_Detalle = Trim(lvFacturacion.ListItems(iLoop).SubItems(3))
                exp_Importe = Trim(lvFacturacion.ListItems(iLoop).SubItems(4))
                iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_Beneficio, False)
                iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_Nombre, False)
                iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_TipoPrestacion, False)
                iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_Detalle, False)
                iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_Importe, False)
                iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, "endline", False)
                iConteoCaracteres = iConteoCaracteres + Len("endline")
                If Not iControlExportacion Then
                    MsgBox "Error en el control de exportación. " & Chr(13) & _
                            "Por favor reinicie el proceso.", vbCritical
                    Exit Sub
                End If
            Next iLoop
        End If
        
        Gstr_FacturacionCerrada = False
        Call Actualizar_Config_INI(7, "Status=open")
        Call Actualizar_Config_INI(8, "File=" & Trim(sNombreArchivoExportacion))
        MsgBox "Archivo " & sNombreArchivoExportacion & " a sido grabado. Recuerde que debe exportarlo y cerrarlo.", vbInformation
        Unload Me
      Else
        MsgBox "error exportando la cabecera", vbInformation
    End If
End Sub



Private Sub cmdExportar_Click()
Dim sNombreArchivoExportacion As String
Dim iElementos As Integer
Dim iLoop As Integer
Dim iControlExportacion As Boolean
'VARIABLES PARA COMPLETAR LA LINEA DE EXPORTACION
Dim exp_RazonSocial As String
Dim exp_Cuit As String
Dim exp_PeriodoFacturado As String
Dim exp_NumeroFactura As String
Dim exp_Beneficio As String
Dim exp_Nombre As String
Dim exp_TipoPrestacion As String
Dim exp_Detalle As String
Dim exp_Importe As String
Dim exp_Archivo As String

    If lvFacturacion.ListItems.Count = 0 Then
        MsgBox "Debe haber al menos un consumo cargado para guardar la facturación del período.", vbInformation
        Exit Sub
    End If
    
    If MsgBox("Al cerrar la facturación la misma queda bloqueada para futuras ediciones." & Chr(13) & Chr(13) & "¿ Desea continuar ?", vbYesNo) = vbNo Then Exit Sub

    exp_EstadoFacturacion = "C" 'SETEAMOS EL ESTADO DE LA FACTURACION SOBRE LA VARIABLE GLOBAL
    sNombreArchivoExportacion = Generar_Archivo_Exportacion(Trim(Gstr_TipoPrestador), Trim(Gstr_PeriodoFacturado), Trim(Gstr_Cuit), Trim(Gstr_NumeroFactura), Trim(exp_EstadoFacturacion))
    iConteoCaracteres = 0
    iElementos = lvFacturacion.ListItems.Count
    'GRABAMOS CABECERA
    iConteoCaracteres = iConteoCaracteres + Len(Trim("[STATUS]"))
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, Trim("[STATUS]"), True)
    iConteoCaracteres = iConteoCaracteres + Len(Trim("SF=closed"))
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, Trim("SF=closed"), False)
    iConteoCaracteres = iConteoCaracteres + Len(Trim("[STATUS]"))
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, Trim("[HEADER]"), False)
    iConteoCaracteres = iConteoCaracteres + Len(Trim("[HEADER]"))
    exp_RazonSocial = "PRESTADOR=" & Gstr_RazonSocial
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_RazonSocial, False)
    iConteoCaracteres = iConteoCaracteres + Len(Trim(exp_RazonSocial))
    exp_Cuit = "CUIT=" & Gstr_Cuit
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_Cuit, False)
    iConteoCaracteres = iConteoCaracteres + Len(Trim(exp_Cuit))
    exp_PeriodoFacturado = "PERIODOFACTURADO=" & Gstr_PeriodoFacturado
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_PeriodoFacturado, False)
    iConteoCaracteres = iConteoCaracteres + Len(Trim(exp_PeriodoFacturado))
    exp_NumeroFactura = "NUMEROFACTURA=" & Gstr_NumeroFactura
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_NumeroFactura, False)
    iConteoCaracteres = iConteoCaracteres + Len(Trim(exp_NumeroFactura))
    iConteoCaracteres = iConteoCaracteres + Len(Trim("[DETAIL]"))
    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, "[DETAIL]", False)
    If iControlExportacion Then
        If iElementos > 0 Then
            For iLoop = 1 To iElementos
                exp_Beneficio = Trim(lvFacturacion.ListItems(iLoop))
                iConteoCaracteres = iConteoCaracteres + Len(exp_Beneficio)
                exp_Nombre = Trim(lvFacturacion.ListItems(iLoop).SubItems(1))
                iConteoCaracteres = iConteoCaracteres + Len(exp_Nombre)
                exp_TipoPrestacion = Trim(lvFacturacion.ListItems(iLoop).SubItems(2))
                iConteoCaracteres = iConteoCaracteres + Len(Trim(exp_TipoPrestacion))
                exp_Detalle = Trim(lvFacturacion.ListItems(iLoop).SubItems(3))
                iConteoCaracteres = iConteoCaracteres + Len(exp_Detalle)
                exp_Importe = Trim(lvFacturacion.ListItems(iLoop).SubItems(4))
                iConteoCaracteres = iConteoCaracteres + Len(exp_Importe)
                
                iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_Beneficio, False)
                iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_Nombre, False)
                iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_TipoPrestacion, False)
                iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_Detalle, False)
                iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, exp_Importe, False)
                iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, "endline", False)
                iConteoCaracteres = iConteoCaracteres + Len("endline")

                If Not iControlExportacion Then
                    'MOSTRAR MENSAJE DE ERROR Y BORRAR EL ARCHIVO
                End If
            Next iLoop
            iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, "[BOTTOM]", False)
            iConteoCaracteres = iConteoCaracteres + Len(Trim("[BOTTOM]"))
            'AGREGAMOS LOS TOTALES
            Select Case Gstr_TipoPrestador
                Case "Discapacidad"
                    auxTotalImprimir = "TotalPrestacion=" & Trim(txtTotalPrestacion.Text)
                    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, auxTotalImprimir, False)
                    iConteoCaracteres = iConteoCaracteres + Len(Trim(auxTotalImprimir))
                    
                    auxTotalImprimir = "TotalTransporte=" & Trim(txtTotalTransporte.Text)
                    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, auxTotalImprimir, False)
                    iConteoCaracteres = iConteoCaracteres + Len(Trim(auxTotalImprimir))
            
                    auxTotalImprimir = "TotalApoyo=" & Trim(txtTotalApoyo.Text)
                    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, auxTotalImprimir, False)
                    iConteoCaracteres = iConteoCaracteres + Len(Trim(auxTotalImprimir))
            
                    auxTotalImprimir = "TotalMatricula=" & Trim(txtTotalMatricula.Text)
                    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, auxTotalImprimir, False)
                    iConteoCaracteres = iConteoCaracteres + Len(Trim(auxTotalImprimir))
                Case "Hemodialisis"
                    auxTotalImprimir = "TotalPrestacion=" & Trim(txtTotalHemodialisis.Text)
                    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, auxTotalImprimir, False)
                    iConteoCaracteres = iConteoCaracteres + Len(Trim(auxTotalImprimir))
                    
                    auxTotalImprimir = "TotalTransporte=" & Trim(txtTotalTransporteHemodialisis.Text)
                    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, auxTotalImprimir, False)
                    iConteoCaracteres = iConteoCaracteres + Len(Trim(auxTotalImprimir))
                Case "Geriatria"
                    auxTotalImprimir = "TotalPrestacion=" & Trim(txtTotalGeriatria.Text)
                    iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, auxTotalImprimir, False)
                    iConteoCaracteres = iConteoCaracteres + Len(Trim(auxTotalImprimir))
            End Select
            
            iConteoCaracteres = iConteoCaracteres + Len(Trim("CheckSum=") & Trim(Str$(iConteoCaracteres)))
            iControlExportacion = Generar_Exportacion(sNombreArchivoExportacion, Trim("CheckSum=") & Trim(Str$(iConteoCaracteres)), False)
        End If
        
        Call Actualizar_Config_INI(7, "Status=closed")
        
        MsgBox "Archivo " & sNombreArchivoExportacion & " a sido exportado correctamente." & Chr(13) & Chr(13) & _
                        "Prepare la impresora", vbInformation
        Call Actualizar_Config_INI(8, "File=null")
        
        'BORRAMOS EL ARCHIVO DE FACTURACION ABIERTA
        If Not Gstr_FacturacionCerrada Then
            BorrarArchivo (App.Path & "\" & Gstr_NombreArchivoFacturacionAbierta & ".txt")
        End If
        Gstr_FacturacionCerrada = True
        'MOVEMOS EL ARCHIVO A LA CARPETA DE FACTURACIONES CERRADAS
        exp_Archivo = sNombreArchivoExportacion & ".txt"
        MoveFile App.Path & "\" & exp_Archivo, App.Path & Gstr_DestinationPath & "\" & sNombreArchivoExportacion & ".txt"
        Select Case Gstr_TipoPrestador
            Case "Discapacidad"
                Call Imprimir_ListView(lvFacturacion, exp_Archivo, Gstr_RazonSocial, Gstr_Cuit, Gstr_PeriodoFacturado, Gstr_NumeroFactura, saux_TotalGeneral, Trim(txtTotalPrestacion.Text), Trim(txtTotalTransporte.Text), Trim(txtTotalApoyo.Text), Trim(txtTotalMatricula.Text))
            Case "Hemodialisis"
                Call Imprimir_ListView(lvFacturacion, exp_Archivo, Gstr_RazonSocial, Gstr_Cuit, Gstr_PeriodoFacturado, Gstr_NumeroFactura, saux_TotalGeneral, Trim(txtTotalHemodialisis.Text), Trim(txtTotalTransporteHemodialisis.Text))
            Case "Geriatria"
                Call Imprimir_ListView(lvFacturacion, exp_Archivo, Gstr_RazonSocial, Gstr_Cuit, Gstr_PeriodoFacturado, Gstr_NumeroFactura, saux_TotalGeneral, Trim(txtTotalGeriatria.Text), Trim(txtTotalGeriatria.Text))
        End Select
        Unload Me
      Else
        MsgBox "Error exportando la cabecera", vbCritical
    End If
End Sub

Sub Cargar_Facturacion()
    Dim sLinea As String
    Dim sRenglon As String
    Dim sArchivo As String
    
    sArchivo = Trim(Gstr_NombreArchivoFacturacionAbierta & ".txt")
    
    Gstr_PeriodoFacturado = ConfigGetEXTENDIDO(sArchivo, "HEADER", "PERIODOFACTURADO", 1)
    Gstr_NumeroFactura = ConfigGetEXTENDIDO(sArchivo, "HEADER", "NUMEROFACTURA", 1)
    Call Configurar_Cabecera
    
    sRenglon = Chr(13) & Chr(10)
    Open App.Path & "\" & sArchivo For Input As #1
    Line Input #1, sLinea: Line Input #1, sLinea
    Line Input #1, sLinea: Line Input #1, sLinea
    Line Input #1, sLinea: Line Input #1, sLinea
    Line Input #1, sLinea: Line Input #1, sLinea

    While Not EOF(1)
    'EN LA LINEA 8 COMIENZA EL REGISTRO
        Line Input #1, sLinea: sBeneficio = sLinea
        Line Input #1, sLinea: sNombre = sLinea
        Line Input #1, sLinea: sPrestacion = sLinea
        Line Input #1, sLinea: sDescripcion = sLinea
        Line Input #1, sLinea: sImporte = sLinea
        Call Agregar_Item(sBeneficio, sNombre, sPrestacion, sDescripcion, sImporte)
        Call Actualizar_Totales(False)
        Line Input #1, sLinea
        If sLinea <> "endline" Then
            MsgBox "Error importando el archivo.", vbCritical
        End If
    Wend
    Close #1
End Sub

'CARGA LA FACTURACION ABIERTA EN PROCESO
Private Sub Cargar_Facturacion_Abierta()
'    bActualizando = True
    Cargar_Facturacion
End Sub

'BORRAMOS EL ARCHIVO DE FACTURACION ABIERTA
Private Sub cmdBorrarFacturacionActual_Click()
    If MsgBox("Va a eliminar el archivo abierto" & Chr(13) & "¿ Desea continuar ?", vbYesNo) = vbNo Then Exit Sub
    If Not Gstr_FacturacionCerrada Then
        BorrarArchivo (App.Path & "\" & Gstr_NombreArchivoFacturacionAbierta & ".txt")
        Call Actualizar_Config_INI(7, "Status=closed")
        MsgBox "Archivo " & sNombreArchivoExportacion & " a sido eliminado." & Chr(13) & "Puede generar una nueva facturacion.", vbInformation
        Call Actualizar_Config_INI(8, "File=null")
      Else
        MsgBox "Facturacion eliminada.", vbInformation
    End If
    Unload Me
End Sub

'EDITA EL ITEM SELECCIONADO
Private Sub cmdEditarItem_Click()
Dim Indice As Integer
    If lvFacturacion.ListItems.Count > 0 Then
        bEditando = True
        Indice = lvFacturacion.SelectedItem.Index
        Call Cargar_Header(Indice)
        lvFacturacion.ListItems.Remove (Indice)
        Call Actualizar_Totales(False)
        Call Habilitar_Formulario(False)
    End If
End Sub

'ELIMINA EL ITEM SELECCINONADO DEL LISTVIEW
Private Sub cmdEliminarItem_Click()
Dim Indice As Integer
    If lvFacturacion.ListItems.Count > 0 Then
        Indice = lvFacturacion.SelectedItem.Index
        lvFacturacion.ListItems.Remove (Indice)
        Call Actualizar_Totales(False)
    End If
End Sub

Sub Cargar_Header(IndiceLV As Integer)
Dim item_Beneficio As String
Dim item_Nombre As String
Dim item_Prestacion As String
Dim item_Detalle As String
Dim item_Importe As String
    item_Beneficio = lvFacturacion.ListItems(IndiceLV)
    item_Nombre = lvFacturacion.ListItems(IndiceLV).SubItems(1)
    item_Prestacion = lvFacturacion.ListItems(IndiceLV).SubItems(2)
    item_Detalle = lvFacturacion.ListItems(IndiceLV).SubItems(3)
    item_Importe = lvFacturacion.ListItems(IndiceLV).SubItems(4)
    txtBeneficio.Text = item_Beneficio
    txtNombre.Text = item_Nombre
    Select Case item_Prestacion
        Case "Prestacion"
            cmbTipoPrestacion.ListIndex = 0
        Case "Transporte"
            cmbTipoPrestacion.ListIndex = 1
        Case "Apoyo"
            cmbTipoPrestacion.ListIndex = 2
        Case "Matricula"
            cmbTipoPrestacion.ListIndex = 3
        Case "Hemodialisis"
            cmbTipoPrestacion.ListIndex = 0
        Case "Geriatria"
            cmbTipoPrestacion.ListIndex = 0
    End Select
    txtDescripcion.Text = item_Detalle
    txtImporte.Text = item_Importe
End Sub

'DESABILITA EL FORMULARIO DURANTE LA EDICION DE ALGUN ITEM DEL LISTVIEW
Sub Habilitar_Formulario(Estado As Boolean)
    If Estado Then
        lvFacturacion.Enabled = True
        cmdGuardar.Enabled = True
        cmdExportar.Enabled = True
        cmdBorrarFacturacionActual.Enabled = True
        cmdEditarItem.Enabled = True
        cmdEliminarItem.Enabled = True
        cmdCargarPlantilla.Enabled = True
      Else
        lvFacturacion.Enabled = False
        cmdGuardar.Enabled = False
        cmdExportar.Enabled = False
        cmdBorrarFacturacionActual.Enabled = False
        cmdEditarItem.Enabled = False
        cmdEliminarItem.Enabled = False
        cmdCargarPlantilla.Enabled = False
    End If
End Sub

'ACTUALIZA LOS TOTALES POR TIPO DE PRESTACION
Sub Actualizar_Totales(Reset As Boolean)
Dim auxArea As String
Dim auxPrestacion As String
Dim dImporte As Double
Dim iElementos As Integer
Dim iLoop As Integer
Dim dTotalPrestacion As Double
Dim dTotalTransporte As Double
Dim dTotalApoyo As Double
Dim dTotalMatricula As Double
Dim dTotalHemodialisis As Double
Dim dTotalGeriatria As Double
Dim auxValorX As String
    dTotalPrestacion = 0
    dTotalTransporte = 0
    dTotalApoyo = 0
    dTotalMatricula = 0
    dTotalHemodialisis = 0
    dTotalGeriatria = 0
    iElementos = lvFacturacion.ListItems.Count
    If Not Reset Then
        For iLoop = 1 To iElementos
            auxPrestacion = lvFacturacion.ListItems(iLoop).SubItems(2)
            dImporte = lvFacturacion.ListItems(iLoop).SubItems(4)
            Select Case Gstr_TipoPrestador
                Case "Discapacidad"
                    Select Case auxPrestacion
                        Case "Prestacion"
                            dTotalPrestacion = dTotalPrestacion + dImporte
                            auxValorX = ConvertirDecimal(dTotalPrestacion)
                            txtTotalPrestacion = Format(Val(auxValorX), "#0.00")
                            sTotalGeneral = Val(ConvertirDecimal(txtTotalPrestacion)) + Val(ConvertirDecimal(txtTotalTransporte)) + Val(ConvertirDecimal(txtTotalApoyo)) + Val(ConvertirDecimal(txtTotalMatricula))
                        Case "Transporte"
                            dTotalTransporte = dTotalTransporte + dImporte
                            auxValorX = ConvertirDecimal(dTotalTransporte)
                            txtTotalTransporte = Format(Val(auxValorX), "#0.00")
                            sTotalGeneral = Val(ConvertirDecimal(txtTotalPrestacion)) + Val(ConvertirDecimal(txtTotalTransporte)) + Val(ConvertirDecimal(txtTotalApoyo)) + Val(ConvertirDecimal(txtTotalMatricula))
                        Case "Apoyo"
                            dTotalApoyo = dTotalApoyo + dImporte
                            auxValorX = ConvertirDecimal(dTotalApoyo)
                            txtTotalApoyo = Format(Val(auxValorX), "#0.00")
                            sTotalGeneral = Val(ConvertirDecimal(txtTotalPrestacion)) + Val(ConvertirDecimal(txtTotalTransporte)) + Val(ConvertirDecimal(txtTotalApoyo)) + Val(ConvertirDecimal(txtTotalMatricula))
                        Case "Matricula"
                            dTotalMatricula = dTotalMatricula + dImporte
                            auxValorX = ConvertirDecimal(dTotalMatricula)
                            txtTotalMatricula = Format(Val(auxValorX), "#0.00")
                            sTotalGeneral = Val(ConvertirDecimal(txtTotalPrestacion)) + Val(ConvertirDecimal(txtTotalTransporte)) + Val(ConvertirDecimal(txtTotalApoyo)) + Val(ConvertirDecimal(txtTotalMatricula))
                    End Select
                Case "Hemodialisis"
                    Select Case auxPrestacion
                        Case "Hemodialisis"
                            dTotalHemodialisis = dTotalHemodialisis + dImporte
                            auxValorX = ConvertirDecimal(dTotalHemodialisis)
                            txtTotalHemodialisis = Format(Val(auxValorX), "#0.00")
                            sTotalGeneral = Val(ConvertirDecimal(txtTotalHemodialisis)) + Val(ConvertirDecimal(txtTotalTransporte))
                        Case "Transporte"
                            dTotalTransporte = dTotalTransporte + dImporte
                            auxValorX = ConvertirDecimal(dTotalTransporte)
                            txtTotalTransporteHemodialisis = Format(Val(auxValorX), "#0.00")
                            sTotalGeneral = Val(ConvertirDecimal(txtTotalHemodialisis)) + Val(ConvertirDecimal(txtTotalTransporte))
                    End Select
                Case "Geriatria"
                    dTotalGeriatria = dTotalGeriatria + dImporte
                    auxValorX = ConvertirDecimal(dTotalGeriatria)
                    txtTotalGeriatria = Format(Val(auxValorX), "#0.00")
                    sTotalGeneral = Val(ConvertirDecimal(txtTotalGeriatria))
            End Select
        Next iLoop
    End If
    saux_TotalGeneral = sTotalGeneral
End Sub


Private Sub cmdCargarPlantilla_Click()
    lvFacturacion.ListItems.Clear
    Call Actualizar_Totales(True)
    Call Habilitar_Formulario(False)
    frmListadoFacturaciones.cmdSeleccionar.Visible = True
    frmListadoFacturaciones.cmdImprimir.Visible = False
    frmListadoFacturaciones.cmdBorrar.Visible = False
    frmListadoFacturaciones.Show vbModal
    If frmListadoFacturaciones.optCancelado Then
        Call Habilitar_Formulario(True)
        Unload frmListadoFacturaciones
        Exit Sub
    End If
    If frmListadoFacturaciones.optSeleccionado Then
        Dim FileFacturacion As String
        Dim Importes As Boolean
            FileFacturacion = frmListadoFacturaciones.fileFacturaciones.FileName
            If MsgBox("¿Importar la plantilla con valores?", vbYesNo) = vbYes Then
 '               Call Habilitar_FrameFacturaciones(False)
                Call Habilitar_Formulario(True)
                Call Cargar_Plantilla(FileFacturacion, True)
              Else
 '               Call Habilitar_FrameFacturaciones(False)
                Call Habilitar_Formulario(True)
                Call Cargar_Plantilla(FileFacturacion, False)
            End If
    End If
    Unload frmListadoFacturaciones
End Sub

'CANCELAMOS LA SELECCION DE UN PERIODO DE FACTURACION
Private Sub cmdCancelar_Click()
'    Call Habilitar_FrameFacturaciones(False)
    Call Habilitar_Formulario(True)
End Sub

Sub Cargar_Plantilla(sFileName As String, CargaImportes As Boolean)
    Dim sLinea As String
    Dim sRenglon As String
    Dim sArchivo As String
    
    sArchivo = Trim(sFileName)
    Call Configurar_Cabecera
    sRenglon = Chr(13) & Chr(10)
    Open App.Path & "\Facturaciones Cerradas\" & sArchivo For Input As #1
    Line Input #1, sLinea: Line Input #1, sLinea
    Line Input #1, sLinea: Line Input #1, sLinea
    Line Input #1, sLinea: Line Input #1, sLinea
    Line Input #1, sLinea: Line Input #1, sLinea

    While Not EOF(1)
    'EN LA LINEA 8 COMIENZA EL REGISTRO
        Line Input #1, sLinea
        If sLinea <> "[BOTTOM]" Then
            sBeneficio = sLinea
          Else
            Close #1
            Exit Sub
        End If
        Line Input #1, sLinea: sNombre = sLinea
        Line Input #1, sLinea: sPrestacion = sLinea
        Line Input #1, sLinea: sDescripcion = sLinea
        Line Input #1, sLinea: sImporte = sLinea
        If CargaImportes Then
            Call Agregar_Item(sBeneficio, sNombre, sPrestacion, sDescripcion, sImporte)
          Else
            Call Agregar_Item(sBeneficio, sNombre, sPrestacion, sDescripcion, "0")
        End If
        Call Actualizar_Totales(False)
        Line Input #1, sLinea
        If sLinea <> "endline" Then
            MsgBox "Error importando el archivo.", vbInformation
        End If
    Wend
    Close #1
    Exit Sub
error:
    Close #1
End Sub

Sub Pasar_a_Mayusculas(Control As TextBox, Estado As Boolean)
    If Estado = True Then
        Control.Text = UCase(Control.Text)
      Else
        Control.Text = LCase(Control.Text)
    End If
End Sub

Private Sub cmdCerrar_Click()
    If MsgBox("Si cierra sin guardar se perderán los cambios realizados." & Chr(13) & "¿ Desea continuar ?", vbYesNo + vbInformation) = vbYes Then
        Unload Me
    End If
End Sub


Private Sub txtNombre_LostFocus()
    Pasar_a_Mayusculas txtNombre, True
End Sub

Private Sub txtDescripcion_LostFocus()
    Pasar_a_Mayusculas txtDescripcion, True
End Sub

Private Sub txtImporte_GotFocus()
    If txtImporte.Text = "0,00" Or txtImporte = "" Then
        txtImporte.Text = ""
      Else
        txtImporte.Text = ConvertirDecimal(txtImporte.Text)
    End If
End Sub
