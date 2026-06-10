VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRepRegistroMarc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Marcaciones"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "FrmRepRegistroMarc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7815
   Begin VB.CheckBox ChkAsistencia 
      Caption         =   "Asistencia por Tipo de Hora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   240
      Width           =   2775
   End
   Begin VB.Frame Frame2 
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
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6765
      Begin VB.Frame FraAsistencia 
         Height          =   855
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   6540
         Begin MSDataListLib.DataCombo dtcTipoHora 
            Height          =   315
            Left            =   1800
            TabIndex        =   22
            Top             =   360
            Width           =   4485
            _ExtentX        =   7911
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Hora"
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
            Left            =   240
            TabIndex        =   21
            Top             =   420
            Width           =   1065
         End
      End
      Begin VB.OptionButton OptDia 
         Caption         =   "Dia"
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
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   855
      End
      Begin VB.OptionButton OptFecha 
         Caption         =   "Rango de Fecha"
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
         Left            =   3480
         TabIndex        =   6
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Frame FraTrabajador 
         Caption         =   "Trabajador "
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
         TabIndex        =   13
         Top             =   1200
         Width           =   6540
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
            TabIndex        =   16
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
            TabIndex        =   15
            Top             =   615
            Width           =   4575
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
            Left            =   5940
            TabIndex        =   14
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
            TabIndex        =   18
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
            TabIndex        =   17
            Top             =   660
            Width           =   1530
         End
      End
      Begin VB.Frame FraFecha 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   3360
         TabIndex        =   8
         Top             =   2280
         Width           =   3255
         Begin MSComCtl2.DTPicker DtpFechaFin 
            Height          =   285
            Left            =   1290
            TabIndex        =   9
            Top             =   720
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
            Format          =   63635457
            CurrentDate     =   38056
         End
         Begin MSComCtl2.DTPicker DtpFechaInicio 
            Height          =   285
            Left            =   1290
            TabIndex        =   10
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
            Format          =   63635457
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
            TabIndex        =   12
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
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   240
         End
      End
      Begin VB.Frame FraDia 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   120
         TabIndex        =   3
         Top             =   2280
         Width           =   3255
         Begin MSComCtl2.DTPicker DTPFecha 
            Height          =   285
            Left            =   1320
            TabIndex        =   4
            Top             =   240
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
            Format          =   63635457
            CurrentDate     =   38056
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
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
            TabIndex        =   5
            Top             =   315
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   2160
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
            Picture         =   "FrmRepRegistroMarc.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRegistroMarc.frx":64EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   1890
      Left            =   6960
      TabIndex        =   1
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
         TabIndex        =   2
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
Attribute VB_Name = "FrmRepRegistroMarc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrformname As String = "Reporte de Marcaciones"

Private Sub Check1_Click()

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

Public Sub ShowForm()
    Me.Show
    CenterForm Me
    Me.Caption = mstrformname
End Sub

Private Sub Form_Load()
Dim objPersonal As Object
On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set Me.dtcTipoHora.RowSource = objPersonal.GetTipoHora()
    Set objPersonal = Nothing
    Me.dtcTipoHora.ListField = "Descripcion"
    Me.dtcTipoHora.BoundColumn = "IdTipoHora"
    Me.DtpFechaFin.Value = Date
    Me.DtpFechaInicio.Value = Date
    Me.DTPFecha.Value = Date
    Me.OptDia.Value = True
    Me.OptFecha.Value = False
    Me.DTPFecha.Enabled = 1
    Me.DtpFechaFin.Enabled = 0
    Me.DtpFechaInicio.Enabled = 0
    
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Call ErrorMessage(FrmRepRegistroMarc_Load, Err.Source & " FrmRepRegistroMarc :Load", Err.Description)
End Sub

Private Sub OptDia_Click()
    Me.DTPFecha.Enabled = 1
    Me.DtpFechaFin.Enabled = 0
    Me.DtpFechaInicio.Enabled = 0
End Sub

Private Sub OptFecha_Click()
    Me.DTPFecha.Enabled = 0
    Me.DtpFechaFin.Enabled = 1
    Me.DtpFechaInicio.Enabled = 1
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
Dim horaInicio As Date
Dim horaFin As Date
Dim intIdTipoHora As Integer

horaInicio = gdatHora_Nula_RangoInicio
horaFin = gdatHora_Nula_RangoFin

If Me.dtcTipoHora.BoundText = "" Then
    intIdTipoHora = 0
Else
    intIdTipoHora = Me.dtcTipoHora.BoundText
End If
    
On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    FrmMensajeEspera.ShowForm
    FrmMensajeEspera.Refresh
    Set objReport = CreateObject("Personal.clsReportes")
    If Me.OptDia.Value Then
        If Me.ChkAsistencia.Value Then
            Set rstDataReport = objReport.GetRegistroMarcacionesDiaTipo(horaInicio, horaFin, Me.DTPFecha.Value, Val(Me.TxtIdTrabajador.Text), intIdTipoHora)
            Set objReport = Nothing
            If RecordsetReporteVacio(rstDataReport, strMensaje) Then
                MsgBox strMensaje, vbCritical, mstrformname
            Else
                gaParametrosReport(3, 1) = "FechaInicio"
                gaParametrosReport(3, 2) = CStr(Me.DTPFecha.Value)
                ShowMultiReport rstDataReport, "rptRepRegistroMarcacionTipoHora", gaParametrosReport()
            End If
        Else
            Set rstDataReport = objReport.GetRegistroMarcacionesDia(horaInicio, horaFin, Me.DTPFecha.Value, Val(Me.TxtIdTrabajador.Text))
            Set objReport = Nothing
            If RecordsetReporteVacio(rstDataReport, strMensaje) Then
                MsgBox strMensaje, vbCritical, mstrformname
            Else
                gaParametrosReport(3, 1) = "FechaInicio"
                gaParametrosReport(3, 2) = CStr(Me.DTPFecha.Value)
                ShowMultiReport rstDataReport, "rptRepRegistroMarcacionDia", gaParametrosReport()
            End If
        End If
    Else
        If Me.ChkAsistencia.Value Then
            Set rstDataReport = objReport.GetRegistroMarcacionesFechaTipo(horaInicio, horaFin, Me.DtpFechaInicio.Value, Me.DtpFechaFin.Value, Val(Me.TxtIdTrabajador.Text), intIdTipoHora)
            Set objReport = Nothing
            If RecordsetReporteVacio(rstDataReport, strMensaje) Then
                MsgBox strMensaje, vbCritical, mstrformname
            Else
                gaParametrosReport(3, 1) = "FechaInicio"
                gaParametrosReport(3, 2) = CStr(Me.DtpFechaInicio.Value)
                gaParametrosReport(4, 1) = "FechaFin"
                gaParametrosReport(4, 2) = CStr(Me.DtpFechaFin.Value)
                ShowMultiReport rstDataReport, "rptRepRegistroMarcacionTipoHora", gaParametrosReport()
            End If
        Else
            Set rstDataReport = objReport.GetRegistroMarcacionesFecha(horaInicio, horaFin, Me.DtpFechaInicio.Value, Me.DtpFechaFin.Value, Val(Me.TxtIdTrabajador.Text))
            Set objReport = Nothing
            If RecordsetReporteVacio(rstDataReport, strMensaje) Then
                MsgBox strMensaje, vbCritical, mstrformname
            Else
                gaParametrosReport(3, 1) = "FechaInicio"
                gaParametrosReport(3, 2) = CStr(Me.DtpFechaInicio.Value)
                gaParametrosReport(4, 1) = "FechaFin"
                gaParametrosReport(4, 2) = CStr(Me.DtpFechaFin.Value)
                ShowMultiReport rstDataReport, "rptRepRegistroMarcacionFecha", gaParametrosReport()
            End If
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
    
Call ErrorMessage(FrmRepRegistroMarc_Imprimir, Err.Source & " FrmRepRegistroMarc:Imprimir", Err.Description)
End Sub

Private Sub Limpiar()
    Me.TxtIdTrabajador.Text = 0
    Me.TxtApellido.Text = ""
    Me.DtpFechaFin.Value = Date
    Me.DtpFechaInicio.Value = Date
    Me.DTPFecha.Value = Date
    Me.OptDia.Value = True
    Me.OptFecha.Value = False
    Me.dtcTipoHora.BoundText = ""
End Sub

