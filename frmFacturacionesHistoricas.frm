VERSION 5.00
Begin VB.Form optSeleccionado 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optCancelado 
      Caption         =   "Option1"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   6240
      Width           =   2775
   End
   Begin VB.OptionButton optSeleccionado 
      Caption         =   "Option1"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Frame frmFacturacionesCerradas 
      BackColor       =   &H80000002&
      Caption         =   "Seleccione una Plantilla de Facturacion"
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
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
         Height          =   4350
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3735
      End
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Seleccionar"
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4920
         Width           =   1575
      End
   End
End
Attribute VB_Name = "optSeleccionado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
