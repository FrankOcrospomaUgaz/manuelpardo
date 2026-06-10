VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBusqCompromisosMarcacion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBusqCompromisosMarcacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7440
   Begin TabDlg.SSTab sstBusqueda 
      Height          =   3150
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5556
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Criterios de Trabajador"
      TabPicture(0)   =   "FrmBusqCompromisosMarcacion.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraOtros"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraTrabajador"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Criterios de Horario"
      TabPicture(1)   =   "FrmBusqCompromisosMarcacion.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraFecha"
      Tab(1).Control(1)=   "FraHora"
      Tab(1).ControlCount=   2
      Begin VB.Frame FraTrabajador 
         Caption         =   "Trabajador"
         Height          =   1035
         Left            =   120
         TabIndex        =   21
         Top             =   390
         Width           =   6030
         Begin VB.TextBox TxtIdTrabajador 
            Height          =   285
            Left            =   1770
            TabIndex        =   2
            Tag             =   "1"
            Text            =   "0"
            Top             =   210
            Width           =   975
         End
         Begin VB.TextBox TxtApellido 
            Height          =   285
            Left            =   1740
            MaxLength       =   60
            TabIndex        =   3
            Top             =   615
            Width           =   4215
         End
         Begin VB.CommandButton CmdBuscarTrabajador 
            Caption         =   "..."
            Height          =   240
            Left            =   5445
            TabIndex        =   1
            ToolTipText     =   "Búsqueda de Trabajador"
            Top             =   255
            Width           =   465
         End
         Begin VB.Label LblIdTrabajador 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   105
            TabIndex        =   23
            Top             =   285
            Width           =   570
         End
         Begin VB.Label LblApellido 
            AutoSize        =   -1  'True
            Caption         =   "Apellido y Nombre"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   660
            Width           =   1530
         End
      End
      Begin VB.Frame FraFecha 
         Caption         =   "         Rango de Fechas"
         Height          =   720
         Left            =   -74715
         TabIndex        =   18
         Top             =   660
         Width           =   5670
         Begin VB.CheckBox ChkFecha 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   165
            TabIndex        =   7
            Top             =   -15
            Width           =   300
         End
         Begin MSComCtl2.DTPicker DtpFechaFin 
            Height          =   285
            Left            =   3870
            TabIndex        =   26
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
            Format          =   50397185
            CurrentDate     =   38056
         End
         Begin MSComCtl2.DTPicker DtpFechaInicio 
            Height          =   285
            Left            =   1290
            TabIndex        =   27
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
            Format          =   50397185
            CurrentDate     =   38056
         End
         Begin VB.Label LblFechaInicio 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   315
            Width           =   465
         End
         Begin VB.Label LblFechaFin 
            AutoSize        =   -1  'True
            Caption         =   "Fin"
            Height          =   195
            Left            =   3300
            TabIndex        =   19
            Top             =   315
            Width           =   240
         End
      End
      Begin VB.Frame FraHora 
         Caption         =   "         Rango de Horas"
         Height          =   720
         Left            =   -74715
         TabIndex        =   15
         Top             =   1665
         Width           =   5670
         Begin VB.CheckBox ChkHora 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   165
            TabIndex        =   8
            Top             =   0
            Width           =   300
         End
         Begin MSComCtl2.DTPicker DtpHoraInicio 
            Height          =   285
            Left            =   1305
            TabIndex        =   9
            Top             =   285
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
            CustomFormat    =   "hh:mm tt"
            Format          =   50397187
            UpDown          =   -1  'True
            CurrentDate     =   38056
         End
         Begin MSComCtl2.DTPicker DtpHoraFin 
            Height          =   285
            Left            =   3855
            TabIndex        =   10
            Top             =   285
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
            CustomFormat    =   "hh:mm tt"
            Format          =   50397187
            UpDown          =   -1  'True
            CurrentDate     =   38056
         End
         Begin VB.Label LblHoraFin 
            AutoSize        =   -1  'True
            Caption         =   "Fin"
            Height          =   195
            Left            =   3300
            TabIndex        =   17
            Top             =   315
            Width           =   240
         End
         Begin VB.Label LblHoraInicio 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   315
            Width           =   465
         End
      End
      Begin VB.Frame FraOtros 
         Caption         =   "Otros"
         Height          =   1575
         Left            =   135
         TabIndex        =   11
         Top             =   1455
         Width           =   5985
         Begin MSDataListLib.DataCombo DtcCategoria 
            Height          =   315
            Left            =   1710
            TabIndex        =   4
            Top             =   255
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   714
            _Version        =   393216
            Style           =   2
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcCargo 
            Height          =   315
            Left            =   1710
            TabIndex        =   5
            Top             =   690
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   714
            _Version        =   393216
            Style           =   2
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcArea 
            Height          =   315
            Left            =   1710
            TabIndex        =   6
            Top             =   1140
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   714
            _Version        =   393216
            Style           =   2
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label LblArea 
            AutoSize        =   -1  'True
            Caption         =   "Area"
            Height          =   195
            Left            =   165
            TabIndex        =   14
            Top             =   1140
            Width           =   405
         End
         Begin VB.Label LblCargo 
            AutoSize        =   -1  'True
            Caption         =   "Cargo"
            Height          =   195
            Left            =   165
            TabIndex        =   13
            Top             =   660
            Width           =   495
         End
         Begin VB.Label LblCategoria 
            AutoSize        =   -1  'True
            Caption         =   "Categoria"
            Height          =   195
            Left            =   165
            TabIndex        =   12
            Top             =   270
            Width           =   825
         End
      End
   End
   Begin MSComctlLib.ImageList ImgLstAcciones 
      Left            =   6600
      Top             =   2475
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
            Picture         =   "FrmBusqCompromisosMarcacion.frx":0902
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBusqCompromisosMarcacion.frx":360C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   2175
      Left            =   6405
      TabIndex        =   24
      Top             =   75
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
         Height          =   1725
         Left            =   30
         TabIndex        =   25
         Top             =   420
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
Attribute VB_Name = "FrmBusqCompromisosMarcacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strModo As enumActions
Public intIdTrabajador As Integer
Public strApellidoNombre As String
Public intIdCategoria As Integer
Public intIdCargo As Integer
Public intIdArea As Integer
Public datFechaInicio As Date
Public datFechaFin As Date
Public datHoraInicio As Date
Public datHoraFin As Date
Public strApellido As String

Private mrstCategoria As Object
Private mrstCargo As Object
Private mrstArea As Object

Public blnStatus As Boolean
Public strMensajeError As String

Private Const mstrformname As String = "Consulta de Compromiso Marcacion por Trabajador"

Public Property Get rstCategoria() As Variant
    Set rstCategoria = mrstCategoria
End Property
Public Property Set rstCategoria(ByVal vNewValue As Variant)
    Set mrstCategoria = vNewValue
    If Not rstCategoria.EOF Then
        Set Me.DtcCategoria.RowSource = rstCategoria
        Me.DtcCategoria.ListField = "Descripcion"
        Me.DtcCategoria.BoundColumn = "IdCategoria"
        Me.DtcCategoria.Enabled = True
    Else
        Me.DtcCategoria.Enabled = False
        Me.DtcCategoria.Text = ""
    End If
End Property

Public Property Get rstCargo() As Variant
    Set rstCargo = mrstCargo
End Property
Public Property Set rstCargo(ByVal vNewValue As Variant)
    Set mrstCargo = vNewValue
    If Not rstCargo.EOF Then
        Set Me.DtcCargo.RowSource = rstCargo
        Me.DtcCargo.ListField = "Descripcion"
        Me.DtcCargo.BoundColumn = "IdCargo"
        Me.DtcCargo.Enabled = True
    Else
        Me.DtcCargo.Enabled = False
        Me.DtcCargo.Text = ""
    End If
End Property

Public Property Get rstArea() As Variant
    Set rstArea = mrstArea
End Property
Public Property Set rstArea(ByVal vNewValue As Variant)
    Set mrstArea = vNewValue
    If Not rstArea.EOF Then
        Set Me.DtcArea.RowSource = rstArea
        Me.DtcArea.ListField = "Descripcion"
        Me.DtcArea.BoundColumn = "IdArea"
        Me.DtcArea.Enabled = True
    Else
        Me.DtcArea.Enabled = False
        Me.DtcArea.Text = ""
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
    Call ErrorMessage(FrmBusqCompromisoMarcacion_ShowForm, Err.Source & " FrmBusqCompromisoMarcacion:ShowForm", Err.Description)
End Sub

Private Sub CmdBuscarTrabajador_Click()
    FrmBusqTrabajador.strModo = Find
    FrmBusqTrabajador.ShowForm

    If Trim(FrmBusqTrabajador.intIdTrabajador) <> 0 Then
        Me.TxtIdTrabajador.Text = FrmBusqTrabajador.intIdTrabajador
        Me.TxtApellido.Text = Trim$(FrmBusqTrabajador.strApellidoNombre)
    End If
    Set FrmBusqTrabajador = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrstCategoria = Nothing
    Set mrstCargo = Nothing
    Set mrstArea = Nothing
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
    Set rstCategoria = objPersonal.GetCategoria()
    Set rstCargo = objPersonal.getcargo()
    Set rstArea = objPersonal.GetArea()
    Set objPersonal = Nothing
    
    Me.DtcCategoria.BoundText = Me.intIdCategoria
    Me.DtcCargo.BoundText = Me.intIdCargo
    Me.DtcArea.BoundText = Me.intIdArea
    Me.TxtIdTrabajador.Text = 0
    Me.TxtApellido.Text = ""
    Me.DtpFechaFin.Value = Date
    Me.DtpFechaInicio.Value = Date
    Me.DtpHoraInicio.Value = Time
    Me.DtpHoraFin.Value = Time
    Me.ChkFecha.Value = 1
    Me.ChkHora.Value = 0
    Me.sstBusqueda.Tab = 0
    Me.TxtApellido.SetFocus
    
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Call ErrorMessage(FrmBusqCompromisoMarcacion_LoadData, Err.Source & " FrmBusqCompromisoMarcacion:LoadData", Err.Description)
End Sub

Public Function ValidateFields() As Boolean
    ValidateFields = False
    intIdTrabajador = Val(Me.TxtIdTrabajador.Text)
    strApellido = Replace(Me.TxtApellido.Text, "'", "''")
    
    If Me.ChkFecha.Value = 1 Then
        Me.datFechaInicio = Me.DtpFechaInicio.Value
        Me.datFechaFin = Me.DtpFechaFin.Value
    Else
        Me.datFechaInicio = gdatFecha_Nula_RangoInicio
        Me.datFechaFin = gdatFecha_Nula_RangoFin
    End If
    
    If Me.ChkHora.Value = 1 Then
        Me.datHoraInicio = Me.DtpHoraInicio.Value
        Me.datHoraFin = Me.DtpHoraFin.Value
    Else
        Me.datHoraInicio = gdatHora_Nula_RangoInicio
        Me.datHoraFin = gdatHora_Nula_RangoFin
    End If
    
    If Me.DtcCategoria.BoundText = "" Then
        intIdCategoria = 0
    Else
        intIdCategoria = Me.DtcCategoria.BoundText
    End If
    
    If Me.DtcCargo.BoundText = "" Then
        intIdCargo = 0
    Else
        intIdCargo = Me.DtcCargo.BoundText
    End If
    
    If Me.DtcArea.BoundText = "" Then
        intIdArea = 0
    Else
        intIdArea = Me.DtcArea.BoundText
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
        FrmConsCompromisoMarcacion.intIdTrabajador = Me.intIdTrabajador
        FrmConsCompromisoMarcacion.strApellido = Me.strApellido
        FrmConsCompromisoMarcacion.intIdCategoria = Me.intIdCategoria
        FrmConsCompromisoMarcacion.intIdCargo = Me.intIdCargo
        FrmConsCompromisoMarcacion.intIdArea = Me.intIdArea
        FrmConsCompromisoMarcacion.datFechaFin = Me.datFechaFin
        FrmConsCompromisoMarcacion.datFechaInicio = Me.datFechaInicio
        FrmConsCompromisoMarcacion.datHoraInicio = Me.datHoraInicio
        FrmConsCompromisoMarcacion.datHoraFin = Me.datHoraFin
        FrmConsCompromisoMarcacion.strModo = Me.strModo
        FrmConsCompromisoMarcacion.ShowForm
        Unload Me
        Screen.MousePointer = vbDefault
    End If

Exit Function
ErrHandler:
   Queries = False
   Screen.MousePointer = vbDefault

Call ErrorMessage(FrmBusqCompromisoMarcacion_Queries, Err.Source & " FrmBusqCompromisoMarcacion:Queries", Err.Description)
End Function

Private Function Limpiar() As Boolean
    Me.DtcCategoria.BoundText = ""
    Me.DtcCargo.BoundText = ""
    Me.DtcArea.BoundText = ""
    Me.TxtIdTrabajador.Text = 0
    Me.TxtApellido.Text = ""
    Me.DtpFechaFin.Value = Date
    Me.DtpFechaInicio.Value = Date
    Me.DtpHoraInicio.Value = Time
    Me.DtpHoraFin.Value = Time
    Me.ChkFecha.Value = 1
    Me.ChkHora.Value = 0
    Me.sstBusqueda.Tab = 0
    Me.CmdBuscarTrabajador.SetFocus
    Limpiar = True
End Function




