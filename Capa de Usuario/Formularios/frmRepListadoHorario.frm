VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepListadoHorario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado De Horarios"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmRepListadoHorario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7455
   Begin VB.Frame Frame1 
      Caption         =   "Criterios de Reporte"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   105
      TabIndex        =   12
      Top             =   90
      Width           =   6285
      Begin VB.OptionButton OptTodos 
         Caption         =   "Mostrar todos los días"
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
         Left            =   3375
         TabIndex        =   2
         Top             =   345
         Width           =   2700
      End
      Begin VB.OptionButton OptSemana 
         Caption         =   "Mostrar sólo una semana"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   2700
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
         Left            =   135
         TabIndex        =   18
         Top             =   2760
         Width           =   6015
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
            Format          =   78905345
            CurrentDate     =   38056
         End
         Begin MSComCtl2.DTPicker DtpFechaInicio 
            Height          =   285
            Left            =   1290
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
            Format          =   78905345
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
            Left            =   240
            TabIndex        =   20
            Top             =   315
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
            Left            =   3705
            TabIndex        =   19
            Top             =   315
            Width           =   240
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
         Left            =   105
         TabIndex        =   15
         Top             =   690
         Width           =   6060
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
            TabIndex        =   3
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
            Left            =   1770
            MaxLength       =   60
            TabIndex        =   4
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
            Left            =   5460
            TabIndex        =   5
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
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   660
            Width           =   1530
         End
      End
      Begin MSDataListLib.DataCombo dtcCargo 
         Height          =   315
         Left            =   1860
         TabIndex        =   7
         Top             =   2310
         Width           =   4260
         _ExtentX        =   7514
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
      Begin MSDataListLib.DataCombo dtcGrupo 
         Height          =   315
         Left            =   1860
         TabIndex        =   6
         Top             =   1830
         Width           =   4260
         _ExtentX        =   7514
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
      Begin VB.Label Label3 
         Caption         =   "Cargo:"
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
         Left            =   225
         TabIndex        =   14
         Top             =   2310
         Width           =   495
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
         Left            =   225
         TabIndex        =   13
         Top             =   1830
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   2145
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
            Picture         =   "frmRepListadoHorario.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepListadoHorario.frx":64EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   1890
      Left            =   6480
      TabIndex        =   0
      Top             =   135
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   3334
      BandCount       =   1
      Orientation     =   1
      _CBWidth        =   795
      _CBHeight       =   1890
      _Version        =   "6.7.8862"
      Caption1        =   "Acciones"
      Child1          =   "Toolbar1"
      MinHeight1      =   735
      Width1          =   1200
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   1770
         Left            =   30
         TabIndex        =   11
         Top             =   375
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   3122
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
Attribute VB_Name = "frmRepListadoHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public datFechaInicio As Date
Public datFechaFin As Date

Private Const mstrformname As String = "Reporte de Horarios"

Public Sub ShowForm()
    Me.Show
    CenterForm Me
    Me.Caption = mstrformname
End Sub

Private Sub Form_Load()
Dim objPersonal As Object
On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set Me.dtcCargo.RowSource = objPersonal.getcargo()
    Set Me.dtcGrupo.RowSource = objPersonal.GetGrupo()
    Set objPersonal = Nothing
    Me.dtcCargo.ListField = "Descripcion"
    Me.dtcCargo.BoundColumn = "IdCargo"
    Me.dtcGrupo.ListField = "Descripcion"
    Me.dtcGrupo.BoundColumn = "IdGrupo"
    Me.DtpFechaFin.Value = Date
    Me.DtpFechaInicio.Value = Date
    Me.OptTodos.Value = True
    
Exit Sub
ErrHandler:
    Call ErrorMessage(FrmRepListadoHorario_Load, Err.Source & " FrmRepListadoHorario: Load", Err.Description)
    Set objPersonal = Nothing
End Sub

Private Sub OptSemana_Click()
Dim intDiaActual As Integer
     intDiaActual = DateTime.WeekDay(Date, vbMonday)
     Me.DtpFechaInicio.Value = Date + (8 - intDiaActual)
     Me.DtpFechaFin.Value = Me.DtpFechaInicio.Value
End Sub

Private Sub OptTodos_Click()
    Me.DtpFechaInicio.Value = Date
    Me.DtpFechaFin.Value = Date
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
Dim intIdCargo As Integer
Dim intIdGrupo As Integer

On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    FrmMensajeEspera.ShowForm
    FrmMensajeEspera.Refresh
    If OptSemana.Value Then
        Me.datFechaInicio = Me.DtpFechaInicio.Value
        Me.DtpFechaFin.Value = Me.datFechaInicio + 6
        Me.datFechaFin = Me.DtpFechaFin.Value
    Else
        If Me.ChkFecha.Value = 1 Then
            Me.datFechaInicio = Me.DtpFechaInicio.Value
            Me.datFechaFin = Me.DtpFechaFin.Value
        Else
            Me.datFechaInicio = gdatFecha_Nula_RangoInicio
            Me.datFechaFin = gdatFecha_Nula_RangoFin
        End If
    End If
    If Me.dtcCargo.BoundText = "" Then
        intIdCargo = 0
    Else
        intIdCargo = Me.dtcCargo.BoundText
    End If
    If Me.dtcGrupo.BoundText = "" Then
        intIdGrupo = 0
    Else
        intIdGrupo = Me.dtcGrupo.BoundText
    End If
    Set objReport = CreateObject("Personal.clsReportes")
    Set rstDataReport = objReport.GetReportHorariosTrab(Val(Me.TxtIdTrabajador.Text), _
                    intIdGrupo, intIdCargo, datFechaInicio, datFechaFin)
    Set objReport = Nothing
    gaParametrosReport(3, 1) = "FechaInicio"
    gaParametrosReport(3, 2) = CStr(DtpFechaInicio.Value)
    gaParametrosReport(4, 1) = "FechaFin"
    'If OptSemana.Value Then
     '   gaParametrosReport(4, 2) = CStr(DtpFechaInicio.Value)
    'Else
        gaParametrosReport(4, 2) = CStr(DtpFechaFin.Value)
    'End If
    
    If RecordsetReporteVacio(rstDataReport, strMensaje) Then
        MsgBox strMensaje, vbCritical, Me.Caption
    Else
        If OptSemana.Value Then
            ShowMultiReport rstDataReport, "rptRepListadoDeHorarioSemana", gaParametrosReport()
        Else
            ShowMultiReport rstDataReport, "rptRepListadoDeHorario", gaParametrosReport()
        End If
    End If
    FrmMensajeEspera.Hide
    Set rstDataReport = Nothing
    Screen.MousePointer = vbDefault

Exit Sub
ErrHandler:
    FrmMensajeEspera.Hide
    Screen.MousePointer = vbDefault
    Set rstDataReport = Nothing
    Set objReport = Nothing
    
Call ErrorMessage(FrmRepListadoHorario_Imprimir, Err.Source & " FrmRepListadoHorario:Imprimir", Err.Description)
End Sub

Private Sub Limpiar()
    Me.TxtIdTrabajador.Text = 0
    Me.TxtApellido.Text = ""
    Me.dtcCargo.BoundText = ""
    Me.dtcGrupo.BoundText = ""
    Me.DtpFechaFin.Value = Date
    Me.DtpFechaInicio.Value = Date
    Me.ChkFecha.Value = 0
    Me.OptTodos.Value = True
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
