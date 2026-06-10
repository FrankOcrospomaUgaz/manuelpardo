VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBusqJustificacion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "FrmBusqJustificacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7200
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
      Height          =   1035
      Left            =   90
      TabIndex        =   18
      Top             =   -15
      Width           =   6135
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
         Left            =   1815
         TabIndex        =   2
         Text            =   "0"
         Top             =   210
         Width           =   975
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
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   3
         Top             =   615
         Width           =   4215
      End
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
         Left            =   5490
         TabIndex        =   1
         ToolTipText     =   "Búsqueda de Trabajador"
         Top             =   255
         Width           =   465
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
         Left            =   105
         TabIndex        =   20
         Top             =   270
         Width           =   570
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
         Left            =   120
         TabIndex        =   19
         Top             =   660
         Width           =   1530
      End
   End
   Begin VB.Frame FraCriterio 
      Caption         =   "Criterio de Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   105
      TabIndex        =   0
      Top             =   1050
      Width           =   6135
      Begin VB.OptionButton OptTodas 
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4680
         TabIndex        =   6
         Top             =   315
         Width           =   1140
      End
      Begin VB.OptionButton OptNoRemunerada 
         Caption         =   "No Remunerada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2265
         TabIndex        =   5
         Top             =   315
         Width           =   1740
      End
      Begin VB.OptionButton OptRemunerada 
         Caption         =   "Remunerada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   225
         TabIndex        =   4
         Top             =   315
         Width           =   1395
      End
      Begin VB.Frame FraMotivo 
         Caption         =   "Motivo de Justificación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   5895
         Begin MSDataListLib.DataCombo DtcMotivoJustificacion 
            Height          =   315
            Left            =   1635
            TabIndex        =   7
            Top             =   285
            Width           =   4170
            _ExtentX        =   7355
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label LblDescripcion 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
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
            Left            =   210
            TabIndex        =   15
            Top             =   315
            Width           =   975
         End
      End
      Begin VB.Frame FraFecha 
         Caption         =   "         Rango de Fechas"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1590
         Width           =   5895
         Begin VB.CheckBox ChkFecha 
            Height          =   240
            Left            =   165
            TabIndex        =   8
            Top             =   -15
            Width           =   300
         End
         Begin MSComCtl2.DTPicker DtpFechaFin 
            Height          =   285
            Left            =   4275
            TabIndex        =   10
            Top             =   270
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
            Format          =   21299201
            CurrentDate     =   38056
         End
         Begin MSComCtl2.DTPicker DtpFechaInicio 
            Height          =   285
            Left            =   1635
            TabIndex        =   9
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
            Format          =   21299201
            CurrentDate     =   38056
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
            Left            =   225
            TabIndex        =   13
            Top             =   300
            Width           =   465
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
            Left            =   3690
            TabIndex        =   12
            Top             =   330
            Width           =   240
         End
      End
   End
   Begin MSComctlLib.ImageList ImgLstAcciones 
      Left            =   6540
      Top             =   2385
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
            Picture         =   "FrmBusqJustificacion.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBusqJustificacion.frx":35D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   2175
      Left            =   6285
      TabIndex        =   16
      Top             =   60
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
         Height          =   570
         Left            =   30
         TabIndex        =   17
         Top             =   345
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1005
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
Attribute VB_Name = "FrmBusqJustificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strModo As enumActions
Public intIdTrabajador As Integer
Public strRemunerada As String
Public intIdMotivoJustificacion As Integer
Public datFechaInicio As Date
Public datFechaFin As Date

Private mrstMotivoJustificacion As Object

Public blnStatus As Boolean
Public strMensajeError As String

Private Const mstrformname As String = "Consulta de Justificacion"

Public Property Get rstMotivoJustificacion() As Variant
    Set rstMotivoJustificacion = mrstMotivoJustificacion
End Property
Public Property Set rstMotivoJustificacion(ByVal vNewValue As Variant)
    Set mrstMotivoJustificacion = vNewValue
    If Not rstMotivoJustificacion.EOF Then
        Set Me.DtcMotivoJustificacion.RowSource = rstMotivoJustificacion
        Me.DtcMotivoJustificacion.ListField = "Descripcion"
        Me.DtcMotivoJustificacion.BoundColumn = "IdMotivoJustificacion"
        Me.DtcMotivoJustificacion.Enabled = True
    Else
        Me.DtcMotivoJustificacion.Enabled = False
        Me.DtcMotivoJustificacion.Text = ""
    End If
End Property

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
    Set mrstMotivoJustificacion = Nothing
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
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set rstMotivoJustificacion = objPersonal.GetMotivoJustificacion()
    Set objPersonal = Nothing
    
    Me.DtcMotivoJustificacion.BoundText = Me.intIdMotivoJustificacion
    Me.TxtIdTrabajador.Text = 0
    Me.ChkFecha.Value = 1
    Me.DtpFechaFin.Value = Date
    Me.DtpFechaInicio.Value = Date
    Me.OptTodas.Value = True
    Me.CmdBuscarTrabajador.SetFocus
    
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Call ErrorMessage(FrmBusqJustificacion_LoadData, Err.Source & " FrmBusqJustificacion:LoadData", Err.Description)
End Sub

Public Function ValidateFields() As Boolean
    ValidateFields = False
    
    intIdTrabajador = Val(Me.TxtIdTrabajador.Text)
    If Me.ChkFecha.Value = 1 Then
        Me.datFechaInicio = Me.DtpFechaInicio.Value
        Me.datFechaFin = Me.DtpFechaFin.Value
    Else
        Me.datFechaInicio = gdatFecha_Nula_RangoInicio
        Me.datFechaFin = gdatFecha_Nula_RangoFin
    End If
    
    If Me.OptRemunerada.Value Then
        strRemunerada = gstrHoraRemunerada
    ElseIf Me.OptNoRemunerada.Value Then
        strRemunerada = gstrHoraNoRemunerada
    ElseIf Me.OptTodas.Value Then
        strRemunerada = ""
    End If
    
    If Me.DtcMotivoJustificacion.BoundText = "" Then
        intIdMotivoJustificacion = 0
    Else
        intIdMotivoJustificacion = Me.DtcMotivoJustificacion.BoundText
    End If
        
    ValidateFields = True
End Function

Private Function Queries() As Boolean

On Error GoTo ErrHandler:
    If Not ValidateFields Then
        Exit Function
    Else
        Screen.MousePointer = vbHourglass
        Queries = True
        FrmConsJustificacion.intIdMotivoJustificacion = Me.intIdMotivoJustificacion
        FrmConsJustificacion.intIdTrabajador = Me.intIdTrabajador
        FrmConsJustificacion.strRemunerada = Me.strRemunerada
        FrmConsJustificacion.datFechaFin = Me.datFechaFin
        FrmConsJustificacion.datFechaInicio = Me.datFechaInicio
        FrmConsJustificacion.strModo = Me.strModo
        FrmConsJustificacion.ShowForm
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
    Me.TxtApellido.Text = ""
    Me.DtcMotivoJustificacion.BoundText = Me.intIdMotivoJustificacion
    Me.ChkFecha.Value = 0
    Me.DtpFechaFin.Value = Date
    Me.DtpFechaInicio.Value = Date
    Me.OptRemunerada.Value = True
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


