VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRepListJust 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Justificaciones"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   Icon            =   "FrmRepListJust.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7260
   Begin VB.Frame Frame3 
      Caption         =   "Ordenar Por "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   3840
      Width           =   6255
      Begin VB.OptionButton OptTascend 
         Caption         =   "Total Ascendente"
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
         Left            =   1680
         TabIndex        =   25
         Top             =   240
         Width           =   1980
      End
      Begin VB.OptionButton OptNombre 
         Caption         =   "Nombres"
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
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptTdescend 
         Caption         =   "Total Descendente"
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
         Left            =   3840
         TabIndex        =   23
         Top             =   240
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Criterios del Reporte"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6240
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   6015
         Begin VB.OptionButton OptListado 
            Caption         =   "Como Listado"
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
            Left            =   360
            TabIndex        =   21
            Top             =   180
            Width           =   2175
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Incluir Detalle"
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
            Left            =   3000
            TabIndex        =   20
            Top             =   120
            Width           =   1860
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
         TabIndex        =   14
         Top             =   2880
         Width           =   6015
         Begin VB.CheckBox ChkFecha 
            Height          =   240
            Left            =   165
            TabIndex        =   5
            Top             =   -15
            Width           =   300
         End
         Begin MSComCtl2.DTPicker DtpFechaFin 
            Height          =   285
            Left            =   4275
            TabIndex        =   7
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
            Format          =   63373313
            CurrentDate     =   38056
         End
         Begin MSComCtl2.DTPicker DtpFechaInicio 
            Height          =   285
            Left            =   1290
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
            Format          =   63373313
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
            Left            =   3705
            TabIndex        =   16
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
            Left            =   240
            TabIndex        =   15
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
         Height          =   1035
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   6015
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
            Left            =   5460
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
            Left            =   1770
            MaxLength       =   60
            TabIndex        =   2
            Top             =   615
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
            Left            =   1785
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
            Left            =   120
            TabIndex        =   13
            Top             =   660
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
            Left            =   105
            TabIndex        =   12
            Top             =   270
            Width           =   570
         End
      End
      Begin MSDataListLib.DataCombo dtcGrupo 
         Height          =   315
         Left            =   1740
         TabIndex        =   4
         Top             =   1920
         Width           =   4245
         _ExtentX        =   7488
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
      Begin MSDataListLib.DataCombo dtcMotivoJustificacion 
         Height          =   315
         Left            =   1740
         TabIndex        =   18
         Top             =   2385
         Width           =   4245
         _ExtentX        =   7488
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
      Begin VB.Label Label1 
         Caption         =   "Motivo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2385
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Grupo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1965
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6480
      Top             =   1920
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
            Picture         =   "FrmRepListJust.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepListJust.frx":64EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   1890
      Left            =   6405
      TabIndex        =   9
      Top             =   120
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   3334
      BandCount       =   1
      Orientation     =   1
      _CBWidth        =   795
      _CBHeight       =   1890
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "Toolbar1"
      MinHeight1      =   735
      Width1          =   1200
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   630
         Left            =   30
         TabIndex        =   10
         Top             =   345
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1111
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Vista Previa"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Limpiar"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmRepListJust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public datFechaInicio As Date
Public datFechaFin As Date

Private Const mstrformname As String = "Reporte de Justificaciones"

Public Sub ShowForm()
    Me.Show
    CenterForm Me
    Me.Caption = mstrformname
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

Private Sub Form_Load()
Dim objPersonal As Object
On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set Me.dtcGrupo.RowSource = objPersonal.GetGrupo()
    Set Me.DtcMotivoJustificacion.RowSource = objPersonal.GetMotivoJustificacion()
    Set objPersonal = Nothing
    Me.dtcGrupo.ListField = "Descripcion"
    Me.dtcGrupo.BoundColumn = "IdGrupo"
    Me.DtcMotivoJustificacion.ListField = "Descripcion"
    Me.DtcMotivoJustificacion.BoundColumn = "IdMotivoJustificacion"
    Me.DtpFechaInicio.Value = Date
    Me.DtpFechaFin.Value = Date
    Me.OptNombre.Value = True
    Me.OptDetalle.Value = True
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Call ErrorMessage(FrmRepListJust_Load, Err.Source & " FrmRepListJust: Load", Err.Description)
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        Imprimir
    Else
        Limpiar
    End If
End Sub

Private Sub Imprimir()
Dim objReport As Object
Dim rstDataReport As Object
Dim strMensaje As String
Dim intIdGrupo As Integer
Dim intIdMotivoJustificacion As Integer

On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    FrmMensajeEspera.ShowForm
    FrmMensajeEspera.Refresh
    If Me.ChkFecha.Value = 1 Then
        Me.datFechaInicio = Me.DtpFechaInicio.Value
        Me.datFechaFin = Me.DtpFechaFin.Value
    Else
        Me.datFechaInicio = gdatFecha_Nula_RangoInicio
        Me.datFechaFin = gdatFecha_Nula_RangoFin
    End If
    
    If Me.dtcGrupo.BoundText = "" Then
        intIdGrupo = 0
    Else
        intIdGrupo = Me.dtcGrupo.BoundText
    End If
    
    If Me.DtcMotivoJustificacion.BoundText = "" Then
        intIdMotivoJustificacion = 0
    Else
        intIdMotivoJustificacion = Me.DtcMotivoJustificacion.BoundText
    End If
    
    Set objReport = CreateObject("Personal.clsReportes")
        
    If Me.OptDetalle.Value Then
        Set rstDataReport = objReport.GetJustificacionTrab(Val(Me.TxtIdTrabajador.Text), _
                        intIdGrupo, intIdMotivoJustificacion, Me.datFechaInicio, Me.datFechaFin)
        Set objReport = Nothing
        If RecordsetReporteVacio(rstDataReport, strMensaje) Then
            MsgBox strMensaje, vbCritical, Me.Caption
        Else
            gaParametrosReport(3, 1) = "FechaInicio"
            gaParametrosReport(3, 2) = CStr(DtpFechaInicio.Value)
            gaParametrosReport(4, 1) = "FechaFin"
            gaParametrosReport(4, 2) = CStr(DtpFechaFin.Value)
            ShowMultiReport rstDataReport, "rptRepListadoDeJustificaciones", gaParametrosReport()
        End If
    Else
        Set rstDataReport = objReport.GetJustificacionTrabL(Val(Me.TxtIdTrabajador.Text), _
                        intIdGrupo, intIdMotivoJustificacion, Me.datFechaInicio, Me.datFechaFin)
        Set objReport = Nothing
        If Me.OptTascend.Value Then
            rstDataReport.Sort = "Incidencia asc"
        ElseIf Me.OptTdescend.Value Then
            rstDataReport.Sort = "Incidencia desc"
        End If
            
        If RecordsetReporteVacio(rstDataReport, strMensaje) Then
            MsgBox strMensaje, vbCritical, Me.Caption
        Else
            gaParametrosReport(3, 1) = "FechaInicio"
            gaParametrosReport(3, 2) = CStr(DtpFechaInicio.Value)
            gaParametrosReport(4, 1) = "FechaFin"
            gaParametrosReport(4, 2) = CStr(DtpFechaFin.Value)
            ShowMultiReport rstDataReport, "rptRepListadoDeJustificacionesL", gaParametrosReport()
        End If
    End If
    FrmMensajeEspera.Hide
    Set rstDataReport = Nothing
    Screen.MousePointer = vbDefault

Exit Sub
ErrHandler:
    FrmMensajeEspera.Hide
    Set rstDataReport = Nothing
    Set objReport = Nothing
    Screen.MousePointer = vbDefault
    
Call ErrorMessage(FrmRepListJust_Imprimir, Err.Source & " FrmRepListJust:Imprimir", Err.Description)
End Sub

Private Sub Limpiar()
    Me.TxtIdTrabajador.Text = 0
    Me.TxtApellido.Text = ""
    Me.dtcGrupo.BoundText = ""
    Me.DtcMotivoJustificacion.BoundText = ""
    Me.DtpFechaInicio.Value = Date
    Me.DtpFechaFin.Value = Date
    Me.ChkFecha.Value = 0
End Sub

