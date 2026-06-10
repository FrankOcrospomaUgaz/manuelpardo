VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBusqMarcacionTrabajador 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   Icon            =   "FrmBusqMarcacionTrabajador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   7290
   Begin VB.Frame FraFecha 
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   105
      TabIndex        =   5
      Top             =   1305
      Width           =   6105
      Begin MSComCtl2.DTPicker DtpFechaFin 
         Height          =   285
         Left            =   4425
         TabIndex        =   6
         Top             =   300
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   23986177
         CurrentDate     =   38056
      End
      Begin MSComCtl2.DTPicker DtpFechaInicio 
         Height          =   285
         Left            =   1755
         TabIndex        =   7
         Top             =   300
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   23986177
         CurrentDate     =   38056
      End
      Begin VB.Label LblFechaFin 
         AutoSize        =   -1  'True
         Caption         =   "Fin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3825
         TabIndex        =   9
         Top             =   315
         Width           =   240
      End
      Begin VB.Label LblFechaInicio 
         AutoSize        =   -1  'True
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   315
         Width           =   465
      End
   End
   Begin VB.Frame FraTrabajador 
      Caption         =   "Trabajador"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   6105
      Begin VB.CommandButton CmdBuscarTrabajador 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5430
         TabIndex        =   3
         ToolTipText     =   "Búsqueda de Trabajador"
         Top             =   255
         Width           =   465
      End
      Begin VB.TextBox TxtApellido 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1725
         MaxLength       =   60
         TabIndex        =   2
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox TxtIdTrabajador 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1725
         TabIndex        =   1
         Text            =   "0"
         Top             =   210
         Width           =   975
      End
      Begin VB.Label LblApellido 
         AutoSize        =   -1  'True
         Caption         =   "Apellido y Nombre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   645
         Width           =   1530
      End
      Begin VB.Label LblIdTrabajador 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   570
      End
   End
   Begin MSComctlLib.ImageList ImgLstAcciones 
      Left            =   6585
      Top             =   1830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBusqMarcacionTrabajador.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBusqMarcacionTrabajador.frx":35D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   2175
      Left            =   6345
      TabIndex        =   10
      Top             =   45
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   3836
      BandCount       =   1
      FixedOrder      =   -1  'True
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   855
      _CBHeight       =   2175
      _Version        =   "6.0.8169"
      BandBackColor1  =   -2147483635
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   795
      Width1          =   2955
      FixedBackground1=   0   'False
      UseCoolbarPicture1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   1710
         Left            =   30
         TabIndex        =   11
         Top             =   345
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   3016
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImgLstAcciones"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Consultar]"
               Object.ToolTipText     =   "Consultar datos"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Limpiar]"
               Object.ToolTipText     =   "Limpiar criterios de consulta"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmBusqMarcacionTrabajador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strModo As enumActions
Public intIdTrabajador As Integer
Public strApellido As String
Public datFechaInicio As Date
Public datFechaFin As Date

Public blnStatus As Boolean
Public strMensajeError As String

Private Const mstrformname As String = "Consulta de Marcaciones"

Public Sub ShowForm()
On Error GoTo ErrHandler
    LoadData
    If Me.strModo = Find Then
        Me.Show vbModal
    End If
    If Me.strModo = Query Then
        Me.Show
    End If
    Me.ZOrder

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmBusqJustificacion_ShowForm, Err.Source & " FrmBusqJustificacion:ShowForm", Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set mrstMotivoJustificacion = Nothing
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case KEY_ACTION_CLEAR
            Limpiar
        Case KEY_ACTION_QUERY
            Queries
    End Select
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.strMensajeError = ""
    Me.Caption = mstrformname
End Sub

Private Sub LoadData()
Dim objPersonal As Object

On Error GoTo ErrHandler
    
    Me.TxtIdTrabajador.Text = 0
    Me.TxtApellido.Text = ""
    Me.DtpFechaFin.Value = Date
    Me.DtpFechaInicio.Value = Date
    Me.CmdBuscarTrabajador.SetFocus

Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Call ErrorMessage(FrmBusqJustificacion_LoadData, Err.Source & " FrmBusqJustificacion:LoadData", Err.Description)
End Sub

Public Function ValidateFields() As Boolean
    intIdTrabajador = Val(Me.TxtIdTrabajador.Text)
    Me.strApellido = Me.TxtApellido.Text
    Me.datFechaInicio = Me.DtpFechaInicio.Value
    Me.datFechaFin = Me.DtpFechaFin.Value
    ValidateFields = True
End Function

Private Function Queries() As Boolean

On Error GoTo ErrHandler:
    If Not ValidateFields Then
        Exit Function
    Else
        Screen.MousePointer = vbHourglass
        Queries = True
        FrmConsMarcaciones.intIdTrabajador = Me.intIdTrabajador
        FrmConsMarcaciones.strApellido = Me.strApellido
        FrmConsMarcaciones.datFechaFin = Me.datFechaFin
        FrmConsMarcaciones.datFechaInicio = Me.datFechaInicio
        FrmConsMarcaciones.strModo = Me.strModo
        FrmConsMarcaciones.ShowForm
        Unload Me
        Screen.MousePointer = vbDefault
    End If

Exit Function
ErrHandler:
   Queries = False
   Screen.MousePointer = vbDefault

Call ErrorMessage(FrmBusqJustificacion_Queries, Err.Source & " FrmBusqJustificacion:Queries", Err.Description)
End Function

Private Function Limpiar() As Boolean
    Me.TxtIdTrabajador.Text = 0
    Me.DtpFechaFin.Value = Date
    Me.DtpFechaInicio.Value = Date
    Me.CmdBuscarTrabajador.SetFocus
    Limpiar = True
End Function

Private Sub CmdBuscarTrabajador_Click()
    FrmBusqTrabajador.strModo = Find
    FrmBusqTrabajador.ShowForm

    If Trim(FrmBusqTrabajador.intIdTrabajador) <> 0 Then
        Me.TxtIdTrabajador.Text = FrmBusqTrabajador.intIdTrabajador
        Me.TxtApellido.Text = Trim$(FrmBusqTrabajador.strApellidoNombre)
    End If
    Set FrmBusqTrabajador = Nothing
End Sub



