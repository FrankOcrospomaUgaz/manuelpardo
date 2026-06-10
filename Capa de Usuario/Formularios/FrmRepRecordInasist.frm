VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRepRecordInasist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record de Inasistencias"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "FrmRepRecordInasist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7410
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6585
      Top             =   2070
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
            Picture         =   "FrmRepRecordInasist.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRecordInasist.frx":64EE
            Key             =   ""
         EndProperty
      EndProperty
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
      Height          =   4320
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6330
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
         Height          =   855
         Left            =   120
         TabIndex        =   20
         Top             =   3240
         Width           =   6015
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
            Top             =   360
            Width           =   1905
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
            TabIndex        =   22
            Top             =   360
            Width           =   1575
         End
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
            TabIndex        =   21
            Top             =   360
            Width           =   2220
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   600
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
            Left            =   3000
            TabIndex        =   19
            Top             =   120
            Width           =   2220
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
            Left            =   360
            TabIndex        =   18
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.CheckBox ChkSoloDias 
         Caption         =   "Considerar Solo Dias"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   2220
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
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   6015
         Begin MSComCtl2.DTPicker DTPFin 
            Height          =   345
            Left            =   4320
            TabIndex        =   6
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
            Format          =   63438849
            CurrentDate     =   38381
         End
         Begin MSComCtl2.DTPicker DTPInicio 
            Height          =   345
            Left            =   1725
            TabIndex        =   5
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
            Format          =   63438849
            CurrentDate     =   38381
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
            TabIndex        =   15
            Top             =   270
            Width           =   495
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
            TabIndex        =   14
            Top             =   300
            Width           =   615
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
         TabIndex        =   10
         Top             =   1080
         Width           =   6060
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
            TabIndex        =   12
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
            TabIndex        =   11
            Top             =   270
            Width           =   570
         End
      End
      Begin MSDataListLib.DataCombo dtcGrupo 
         Height          =   315
         Left            =   1860
         TabIndex        =   4
         Top             =   2175
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
         Top             =   2220
         Width           =   615
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   1890
      Left            =   6510
      TabIndex        =   8
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
Attribute VB_Name = "FrmRepRecordInasist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrformname As String = "Reporte de Inasistencias"

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
    Me.OptNombre.Value = True
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Call ErrorMessage(FrmRepRecordInasistencias_Load, Err.Source & " FrmRepRecordInasistencias: Load", Err.Description)
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
    If Me.ChkSoloDias Then
        'Imprime el reporte como detalle
        If Me.OptDetalle.Value Then
            Set rstDataReport = objReport.GetRecordInasistenciasSoliDiaD(Val(Me.TxtIdTrabajador.Text), intIdGrupo, Me.DTPInicio.Value, Me.DTPFin.Value)
            Set objReport = Nothing
            If RecordsetReporteVacio(rstDataReport, strMensaje) Then
                MsgBox strMensaje, vbCritical, Me.Caption
            Else
                gaParametrosReport(3, 1) = "FechaInicio"
                gaParametrosReport(3, 2) = CStr(Me.DTPInicio.Value)
                gaParametrosReport(4, 1) = "FechaFin"
                gaParametrosReport(4, 2) = CStr(Me.DTPFin.Value)
                ShowMultiReport rstDataReport, "rptRepRecordInasistenciasD", gaParametrosReport()
            End If
        'Imprime el reporte como listado
        Else
            Set rstDataReport = objReport.GetRecordInasistenciasSoloDiaL(Val(Me.TxtIdTrabajador.Text), intIdGrupo, Me.DTPInicio.Value, Me.DTPFin.Value)
            Set objReport = Nothing
            If Me.OptTascend.Value Then
                rstDataReport.Sort = "TotalInasistencias asc"
            ElseIf Me.OptTdescend.Value Then
                rstDataReport.Sort = "TotalInasistencias desc"
            End If
            
            If RecordsetReporteVacio(rstDataReport, strMensaje) Then
                MsgBox strMensaje, vbCritical, Me.Caption
            Else
                gaParametrosReport(3, 1) = "FechaInicio"
                gaParametrosReport(3, 2) = CStr(Me.DTPInicio.Value)
                gaParametrosReport(4, 1) = "FechaFin"
                gaParametrosReport(4, 2) = CStr(Me.DTPFin.Value)
                ShowMultiReport rstDataReport, "rptRepRecordInasistenciasL", gaParametrosReport()
            End If
        End If
    Else
        'Imprime el reporte como detalle
        If Me.OptDetalle.Value Then
            Set rstDataReport = objReport.GetRecordInasistenciasD(Val(Me.TxtIdTrabajador.Text), intIdGrupo, Me.DTPInicio.Value, Me.DTPFin.Value)
            Set objReport = Nothing
            If RecordsetReporteVacio(rstDataReport, strMensaje) Then
                MsgBox strMensaje, vbCritical, Me.Caption
            Else
                gaParametrosReport(3, 1) = "FechaInicio"
                gaParametrosReport(3, 2) = CStr(Me.DTPInicio.Value)
                gaParametrosReport(4, 1) = "FechaFin"
                gaParametrosReport(4, 2) = CStr(Me.DTPFin.Value)
                ShowMultiReport rstDataReport, "rptRepRecordInasistenciasD", gaParametrosReport()
            End If
        'Imprime el reporte como listado
        Else
            Set rstDataReport = objReport.GetRecordInasistenciasL(Val(Me.TxtIdTrabajador.Text), intIdGrupo, Me.DTPInicio.Value, Me.DTPFin.Value)
            Set objReport = Nothing
            If Me.OptTascend.Value Then
                rstDataReport.Sort = "TotalInasistencias asc"
            ElseIf Me.OptTdescend.Value Then
                rstDataReport.Sort = "TotalInasistencias desc"
            End If
            If RecordsetReporteVacio(rstDataReport, strMensaje) Then
                MsgBox strMensaje, vbCritical, Me.Caption
            Else
                gaParametrosReport(3, 1) = "FechaInicio"
                gaParametrosReport(3, 2) = CStr(Me.DTPInicio.Value)
                gaParametrosReport(4, 1) = "FechaFin"
                gaParametrosReport(4, 2) = CStr(Me.DTPFin.Value)
                ShowMultiReport rstDataReport, "rptRepRecordInasistenciasL", gaParametrosReport()
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
    
Call ErrorMessage(FrmRepRecordInasistencias_Imprimir, Err.Source & " FrmRepRecordInasistencias: Imprimir", Err.Description)
End Sub

Private Sub Limpiar()
    Me.TxtIdTrabajador.Text = 0
    Me.TxtApellido.Text = ""
    Me.dtcGrupo.BoundText = 0
    Me.DTPInicio.Value = Date
    Me.DTPFin.Value = Date
    Me.OptDetalle.Value = True
End Sub


