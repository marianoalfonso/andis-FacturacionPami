VERSION 5.00
Begin VB.Form frmConfiguracionInicial 
   Appearance      =   0  'Flat
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "CONFIGURACION INICIAL"
   ClientHeight    =   3750
   ClientLeft      =   2295
   ClientTop       =   2595
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Configuración de Datos Iniciales"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   4320
         TabIndex        =   9
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox cmbProvincia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   4575
      End
      Begin VB.ComboBox cmbTipoPrestacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1560
         Width           =   4575
      End
      Begin VB.TextBox txtRazonSocial 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox txtCuit 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label lblProvincia 
         AutoSize        =   -1  'True
         Caption         =   "Provincia:"
         Height          =   195
         Left            =   1080
         TabIndex        =   8
         Top             =   2160
         Width           =   705
      End
      Begin VB.Label lblTipoPrestacion 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Prestación:"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Label lblRazonSocial 
         AutoSize        =   -1  'True
         Caption         =   "Razón Social:"
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label lblCuit 
         AutoSize        =   -1  'True
         Caption         =   "Cuit:"
         Height          =   195
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmConfiguracionInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Top = 100
    Me.Left = 100
    Call Cargar_Formulario
End Sub

Sub Cargar_Formulario()
    Setear_Configuracion
    txtCuit.Text = Gstr_Cuit
    txtRazonSocial.Text = Gstr_RazonSocial
    FillCombo_TipoPrestacion
    FillCombo_Provincia
End Sub

Private Sub cmdGrabar_Click()
    If Validar_Formulario Then
    'para llamar a la funcion de modificacion del txt de configuracion inicial, LINEA + VALOR
        Call Actualizar_Config_INI(2, "Cuit=" & Trim(Me.txtCuit.Text))
        Call Actualizar_Config_INI(3, "RazonSocial=" & Trim(Me.txtRazonSocial.Text))
        Call Actualizar_Config_INI(4, "TipoPrestador=" & Me.cmbTipoPrestacion.Text)
        Call Actualizar_Config_INI(5, "NombreProvincia=" & Me.cmbProvincia.Text)
        Call Setear_Configuracion
        Unload Me
      Else
        MsgBox "Debe completar o corregir los datos indicados.", vbInformation
    End If
End Sub

Sub FillCombo_TipoPrestacion()
    cmbTipoPrestacion.AddItem "Discapacidad"
    cmbTipoPrestacion.AddItem "Hemodialisis"
    cmbTipoPrestacion.AddItem "Geriatria"
    If Gstr_TipoPrestador = "Discapacidad" Then cmbTipoPrestacion.ListIndex = 0
    If Gstr_TipoPrestador = "Hemodialisis" Then cmbTipoPrestacion.ListIndex = 1
    If Gstr_TipoPrestador = "Geriatria" Then cmbTipoPrestacion.ListIndex = 2
    If Gstr_TipoPrestador = "pendiente" Then cmbTipoPrestacion.ListIndex = -1: cmbTipoPrestacion.BackColor = vbRed
End Sub

Sub FillCombo_Provincia()
    cmbProvincia.AddItem "Ciudad Autonoma de Buenos Aires"
    cmbProvincia.AddItem "Buenos Aires"
    cmbProvincia.AddItem "Catamarca"
    cmbProvincia.AddItem "Cordoba"
    cmbProvincia.AddItem "Corrientes"
    cmbProvincia.AddItem "Chaco"
    cmbProvincia.AddItem "Chubut"
    cmbProvincia.AddItem "Entre Rios"
    cmbProvincia.AddItem "Formosa"
    cmbProvincia.AddItem "Jujuy"
    cmbProvincia.AddItem "La Pampa"
    cmbProvincia.AddItem "La Rioja"
    cmbProvincia.AddItem "Mendoza"
    cmbProvincia.AddItem "Misiones"
    cmbProvincia.AddItem "Neuquen"
    cmbProvincia.AddItem "Rio Negro"
    cmbProvincia.AddItem "Salta"
    cmbProvincia.AddItem "San Juan"
    cmbProvincia.AddItem "San Luis"
    cmbProvincia.AddItem "Santa Cruz"
    cmbProvincia.AddItem "Santa Fe"
    cmbProvincia.AddItem "Santiago del Estero"
    cmbProvincia.AddItem "Tierra del Fuego"
    cmbProvincia.AddItem "Tucuman"

    If Gstr_Provincia = "Ciudad Autonoma de Buenos Aires" Then cmbProvincia.ListIndex = 0
    If Gstr_Provincia = "Buenos Aires" Then cmbProvincia.ListIndex = 1
    If Gstr_Provincia = "Catamarca" Then cmbProvincia.ListIndex = 2
    If Gstr_Provincia = "Cordoba" Then cmbProvincia.ListIndex = 3
    If Gstr_Provincia = "Corrientes" Then cmbProvincia.ListIndex = 4
    If Gstr_Provincia = "Chaco" Then cmbProvincia.ListIndex = 5
    If Gstr_Provincia = "Chubut" Then cmbProvincia.ListIndex = 6
    If Gstr_Provincia = "Entre Rios" Then cmbProvincia.ListIndex = 7
    If Gstr_Provincia = "Formosa" Then cmbProvincia.ListIndex = 8
    If Gstr_Provincia = "Jujuy" Then cmbProvincia.ListIndex = 9
    If Gstr_Provincia = "La Pampa" Then cmbProvincia.ListIndex = 10
    If Gstr_Provincia = "La Rioja" Then cmbProvincia.ListIndex = 11
    If Gstr_Provincia = "Mendoza" Then cmbProvincia.ListIndex = 12
    If Gstr_Provincia = "Misiones" Then cmbProvincia.ListIndex = 13
    If Gstr_Provincia = "Neuquen" Then cmbProvincia.ListIndex = 14
    If Gstr_Provincia = "Rio Negro" Then cmbProvincia.ListIndex = 15
    If Gstr_Provincia = "Salta" Then cmbProvincia.ListIndex = 16
    If Gstr_Provincia = "San Juan" Then cmbProvincia.ListIndex = 17
    If Gstr_Provincia = "San Luis" Then cmbProvincia.ListIndex = 18
    If Gstr_Provincia = "Santa Cruz" Then cmbProvincia.ListIndex = 19
    If Gstr_Provincia = "Santa Fe" Then cmbProvincia.ListIndex = 20
    If Gstr_Provincia = "Santiago del Estero" Then cmbProvincia.ListIndex = 21
    If Gstr_Provincia = "Tierra del Fuego" Then cmbProvincia.ListIndex = 22
    If Gstr_Provincia = "Tucuman" Then cmbProvincia.ListIndex = 23
    
    If Gstr_Provincia = "pendiente" Then cmbProvincia.ListIndex = -1: cmbProvincia.BackColor = vbRed
End Sub


Function Validar_Formulario() As Boolean
Dim Cuit_Valido As Boolean
Dim RazonSocial_Valido As Boolean
Dim TipoPrestador_Valido As Boolean
Dim Provincia_Valido As Boolean
    Validar_Formulario = True
    Cuit_Valido = Validar_Cuit(Trim(Me.txtCuit.Text))
    RazonSocial_Valido = Validar_TextBoxNulo(Trim(Me.txtRazonSocial.Text))
    Provincia_Valido = Validar_TextBoxNulo(Trim(Me.cmbProvincia.Text))
    TipoPrestador_Valido = Validar_TextBoxNulo(Trim(Me.cmbTipoPrestacion.Text))
    
    'analizamos cada evaluacion
    If Not Cuit_Valido Then
        Me.txtCuit.BackColor = vbRed
        Validar_Formulario = False
      Else
        Me.txtCuit.BackColor = vbWhite
    End If
    If Not RazonSocial_Valido Then
        Me.txtRazonSocial.BackColor = vbRed
        Validar_Formulario = False
      Else
        Me.txtRazonSocial.BackColor = vbWhite
    End If
    If Not Provincia_Valido Then
        Me.cmbProvincia.BackColor = vbRed
        Validar_Formulario = False
      Else
        Me.cmbProvincia.BackColor = vbWhite
    End If
    If Not TipoPrestador_Valido Then
        Me.cmbTipoPrestacion.BackColor = vbRed
        Validar_Formulario = False
      Else
        Me.cmbTipoPrestacion.BackColor = vbWhite
    End If
               
End Function




