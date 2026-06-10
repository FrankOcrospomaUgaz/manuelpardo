VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRepRecordSobretiempo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record De Sobretiempos"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmRepRecordSobretiempo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7410
   Begin VB.Frame Frame2 
      Caption         =   "Criterios de Reporte"
      Height          =   3105
      Left            =   15
      TabIndex        =   10
      Top             =   0
      Width           =   6330
      Begin VB.OptionButton OptDetalle 
         Caption         =   "Incluir Detalle"
         Height          =   345
         Left            =   3255
         TabIndex        =   2
         Top             =   300
         Width           =   2940
      End
      Begin VB.OptionButton OptListado 
         Caption         =   "Como Listado"
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   1950
      End
      Begin VB.Frame FraTrabajador 
         Caption         =   "Trabajador"
         Height          =   1035
         Left            =   150
         TabIndex        =   14
         Top             =   705
         Width           =   6060
         Begin VB.TextBox TxtIdTrabajador 
            Height          =   285
            Left            =   1785
            TabIndex        =   3
            Text            =   "0"
            Top             =   210
            Width           =   975
         End
         Begin VB.TextBox TxtApellido 
            Height          =   285
            Left            =   1770
            MaxLength       =   60
            TabIndex        =   4
            Top             =   615
            Width           =   4215
         End
         Begin VB.CommandButton CmdBuscarTrabajador 
            Caption         =   "..."
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
            Height          =   195
            Left            =   105
            TabIndex        =   16
            Top             =   270
            Width           =   570
         End
         Begin VB.Label LblApellido 
            AutoSize        =   -1  'True
            Caption         =   "Apellido y Nombre"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   660
            Width           =   1530
         End
      End
      Begin VB.Frame FraFechas 
         Caption         =   "Rango de Fechas"
         Height          =   690
         Left            =   180
         TabIndex        =   11
         Top             =   2295
         Width           =   6015
         Begin MSComCtl2.DTPicker DTPFin 
            Height          =   345
            Left            =   4320
            TabIndex        =   8
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
            Format          =   21364737
            CurrentDate     =   38381
         End
         Begin MSComCtl2.DTPicker DTPInicio 
            Height          =   345
            Left            =   1725
            TabIndex        =   7
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
            Format          =   21364737
            CurrentDate     =   38381
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   255
            Left            =   165
            TabIndex        =   13
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta:"
            Height          =   255
            Left            =   3630
            TabIndex        =   12
            Top             =   270
            Width           =   495
         End
      End
      Begin MSDataListLib.DataCombo dtcGrupo 
         Height          =   315
         Left            =   1905
         TabIndex        =   6
         Top             =   1845
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
         Caption         =   "Grupo"
         Height          =   255
         Left            =   285
         TabIndex        =   17
         Top             =   1890
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6585
      Top             =   1995
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
            Picture         =   "FrmRepRecordSobretiempo.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRecordSobretiempo.frx":64EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   1890
      Left            =   6420
      TabIndex        =   0
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
      Child1          =   "Toolbar1"
      MinHeight1      =   735
      Width1          =   1200
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   630
         Left            =   30
         TabIndex        =   9
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
Attribute VB_Name = "FrmRepRecordSobretiempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrformname As String = "Reporte de Sobretiempos"

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
    Set objPersonal = Nothing
    Me.dtcGrupo.ListField = "Descripcion"
    Me.dtcGrupo.BoundColumn = "IdGrupo"
    Me.DTPInicio.Value = Date
    Me.DTPFin.Value = Date
    Me.OptDetalle.Value = True

Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Call ErrorMessage(FrmRepRecordSobretiempos_Load, Err.Source & " FrmRepRecordSobretiempos: Load", Err.Description)
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

On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    FrmMensajeEspera.ShowForm
    FrmMensajeEspera.Refresh
    If Me.dtcGrupo.BoundText = "" Then
        intIdGrupo = 0
    Else
        intIdGrupo = Me.dtcGrupo.BoundText
    End If
    
    Set objReport = CreateObject("Personal.clsReportes")
    'Imprime el reporte como detalle
    If Me.OptDetalle.Value Then
        Set rstDataReport = objReport.GetRecordSobretiemposD(Val(Me.TxtIdTrabajador.Text), intIdGrupo, Me.DTPInicio.Value, Me.DTPFin.Value)
        Set objReport = Nothing
        If RecordsetReporteVacio(rstDataReport, strMensaje) Then
            MsgBox strMensaje, vbCritical, Me.Caption
        Else
            gaParametrosReport(3, 1) = "FechaInicio"
            gaParametrosReport(3, 2) = CStr(Me.DTPInicio.Value)
            gaParametrosReport(4, 1) = "FechaFin"
            gaParametrosReport(4, 2) = CStr(Me.DTPFin.Value)
            ShowMultiReport rstDataReport, "rptRepRecordDeSobretiemposD", gaParametrosReport()
        End If
    'Imprime el reporte como listado
    Else
        Set rstDataReport = objReport.GetRecordSobretiemposL(Val(Me.TxtIdTrabajador.Text), intIdGrupo, Me.DTPInicio.Value, Me.DTPFin.Value)
        Set objReport = Nothing
        If RecordsetReporteVacio(rstDataReport, strMensaje) Then
            MsgBox strMensaje, vbCritical, Me.Caption
        Else
            gaParametrosReport(3, 1) = "FechaInicio"
            gaParametrosReport(3, 2) = CStr(Me.DTPInicio.Value)
            gaParametrosReport(4, 1) = "FechaFin"
            gaParametrosReport(4, 2) = CStr(Me.DTPFin.Value)
            ShowMultiReport rstDataReport, "rptRepRecordDeSobretiemposL", gaParametrosReport()
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
Call ErrorMessage(FrmRepRecordSobretiempos_Imprimir, Err.Source & " FrmRepRecordSobretiempos:Imprimir", Err.Description)
End Sub

Private Sub Limpiar()
    Me.TxtIdTrabajador.Text = 0
    Me.TxtApellido.Text = ""
    Me.dtcGrupo.BoundText = 0
    Me.DTPInicio.Value = Date
    Me.DTPFin.Value = Date
    Me.OptDetalle.Value = True
End Sub

