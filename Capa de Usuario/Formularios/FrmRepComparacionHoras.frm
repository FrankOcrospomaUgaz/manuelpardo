VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRepComparacionHoras 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "FrmRepComparacionHoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7410
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
      Height          =   3825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6450
      Begin VB.Frame Frame3 
         Caption         =   "Criterios de Listado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   6060
         Begin VB.CheckBox ChkAsistencia 
            Caption         =   "Relación de asistentes"
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
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   2415
         End
         Begin MSDataListLib.DataCombo dtcGrupo 
            Height          =   315
            Left            =   1800
            TabIndex        =   20
            Top             =   600
            Width           =   4095
            _ExtentX        =   7223
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
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   720
            Width           =   510
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   6015
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
            Left            =   3240
            TabIndex        =   16
            Top             =   120
            Width           =   2100
         End
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
            Left            =   240
            TabIndex        =   15
            Top             =   120
            Width           =   2295
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
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   6060
         Begin VB.TextBox TxtIdTrabajador 
            Enabled         =   0   'False
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
            TabIndex        =   9
            Text            =   "0"
            Top             =   210
            Width           =   975
         End
         Begin VB.TextBox TxtApellido 
            Enabled         =   0   'False
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
            TabIndex        =   8
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
            TabIndex        =   7
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
            TabIndex        =   11
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
            TabIndex        =   10
            Top             =   660
            Width           =   1530
         End
      End
      Begin VB.Frame FraFechas 
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
         Height          =   690
         Left            =   240
         TabIndex        =   1
         Top             =   3000
         Width           =   6015
         Begin MSComCtl2.DTPicker DTPFin 
            Height          =   345
            Left            =   4305
            TabIndex        =   2
            Top             =   210
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   609
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
            CurrentDate     =   38381
         End
         Begin MSComCtl2.DTPicker DTPInicio 
            Height          =   345
            Left            =   1725
            TabIndex        =   3
            Top             =   210
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   609
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
            CurrentDate     =   38381
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
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
            Left            =   165
            TabIndex        =   5
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta:"
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
            Left            =   3630
            TabIndex        =   4
            Top             =   270
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   1800
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
            Picture         =   "FrmRepComparacionHoras.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepComparacionHoras.frx":64EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   1890
      Left            =   6540
      TabIndex        =   12
      Top             =   75
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   3334
      BandCount       =   1
      Orientation     =   1
      _CBWidth        =   795
      _CBHeight       =   1890
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   735
      Width1          =   1200
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   630
         Left            =   30
         TabIndex        =   13
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
               Key             =   "[Imprimir]"
               Object.ToolTipText     =   "Ver Vista Previa"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Limpiar]"
               Object.ToolTipText     =   "Limpiar"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmRepComparacionHoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrformname As String = "Reporte de Comparación Horas Trabajadas"
Public Sub ShowForm()
    Me.Show
    CenterForm Me
    Me.Caption = mstrformname
End Sub

Private Sub ChkAsistencia_Click()
    If Me.ChkAsistencia.Value Then
        Me.dtcGrupo.Enabled = 1
    Else
        Me.dtcGrupo.Enabled = 0
    End If
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
    Set objPersonal = Nothing
    Me.dtcGrupo.ListField = "Descripcion"
    Me.dtcGrupo.BoundColumn = "IdGrupo"
    Me.DTPInicio.Value = Date
    Me.DTPFin.Value = Date
    Me.Frame3.Enabled = 0
    Me.dtcGrupo.Enabled = 0
    Me.OptDetalle.Value = 1
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
        Call ErrorMessage(FrmRepComparacionHoras_Load, Err.Source & " FrmRepComparacionHoras: Load", Err.Description)
End Sub

Private Sub OptDetalle_Click()
    If Me.OptDetalle.Value Then
        Me.Frame3.Enabled = 0
    End If
End Sub

Private Sub OptListado_Click()
    If Me.OptListado.Value Then
        Me.Frame3.Enabled = 1
    End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_ACTION_PRINT
        Imprimir
    Case KEY_ACTION_CLEAR
        Limpiar
End Select
End Sub

Private Sub Imprimir()
Dim objReporte As Object
Dim rstReporte As Object
Dim intIdGrupo As Integer

On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    FrmMensajeEspera.ShowForm
    FrmMensajeEspera.Refresh
        
    If Me.dtcGrupo.BoundText = "" Then
        intIdGrupo = 0
    Else
        intIdGrupo = Me.dtcGrupo.BoundText
    End If
    Set objReporte = CreateObject("Personal.clsReportes")
    '** Imprime como detalle
    If Me.OptDetalle.Value Then
        Set rstReporte = objReporte.GetComparacionHoraTrab(Me.DTPInicio.Value, _
                                                        Me.DTPFin.Value, Val(Me.TxtIdTrabajador.Text))
        Set objReporte = Nothing
        If rstReporte.RecordCount = 0 Then
            MsgBox MSG_INFORM_NOTEXISTREGISTEREPORT, vbCritical, mstrformname
        Else
            gaParametrosReport(3, 1) = "FechaInicio"
            gaParametrosReport(3, 2) = CStr(Me.DTPInicio.Value)
            gaParametrosReport(4, 1) = "FechaFin"
            gaParametrosReport(4, 2) = CStr(Me.DTPFin.Value)
            ShowMultiReport rstReporte, "rptRepComparacionHoraTrab", gaParametrosReport
        End If
    Else
    '** Imprime como listado
        If Me.ChkAsistencia.Value Then
            Set rstReporte = objReporte.GetComparacionHoraTrabAsistencia(Me.DTPInicio.Value, _
                                                Me.DTPFin.Value, intIdGrupo, Val(Me.TxtIdTrabajador.Text))
            Set objReporte = Nothing
            If rstReporte.RecordCount = 0 Then
                MsgBox MSG_INFORM_NOTEXISTREGISTEREPORT, vbCritical, mstrformname
            Else
                gaParametrosReport(3, 1) = "FechaInicio"
                gaParametrosReport(3, 2) = CStr(Me.DTPInicio.Value)
                gaParametrosReport(4, 1) = "FechaFin"
                gaParametrosReport(4, 2) = CStr(Me.DTPFin.Value)
                ShowMultiReport rstReporte, "rptRepComparacionHoraTrabAsistencia", gaParametrosReport
            End If
        Else
            Set rstReporte = objReporte.GetComparacionHoraTrabL(Me.DTPInicio.Value, _
                                                Me.DTPFin.Value, Val(Me.TxtIdTrabajador.Text))
            Set objReporte = Nothing
            If rstReporte.RecordCount = 0 Then
                MsgBox MSG_INFORM_NOTEXISTREGISTEREPORT, vbCritical, mstrformname
            Else
                gaParametrosReport(3, 1) = "FechaInicio"
                gaParametrosReport(3, 2) = CStr(Me.DTPInicio.Value)
                gaParametrosReport(4, 1) = "FechaFin"
                gaParametrosReport(4, 2) = CStr(Me.DTPFin.Value)
                ShowMultiReport rstReporte, "rptRepComparacionHoraTrabL", gaParametrosReport
            End If
        End If
    End If
    FrmMensajeEspera.Hide
    Set rstReporte = Nothing
    Screen.MousePointer = vbDefault

Exit Sub
ErrHandler:
    FrmMensajeEspera.Hide
    Set objReporte = Nothing
    Set rstReporte = Nothing
    Screen.MousePointer = vbDefault
    Call ErrorMessage(FrmRepComparacionHoras_Imprimir, Err.Source & " FrmRepDetAsisTrab:Imprimir", Err.Description)
End Sub

Private Sub Limpiar()
    Me.TxtIdTrabajador.Text = 0
    Me.TxtApellido.Text = ""
    Me.dtcGrupo.BoundText = ""
    Me.DTPInicio.Value = Date
    Me.DTPFin.Value = Date
    Me.OptDetalle.Value = 1
End Sub


