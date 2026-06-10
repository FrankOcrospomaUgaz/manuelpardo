VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRepConsolidadoxTrabajador 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   Icon            =   "FrmRepConsolidadoxTrabajador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   7890
   Begin VB.Frame Frame1 
      Caption         =   "Rango de fechas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   6765
      Begin MSComCtl2.DTPicker DTPFechaInicio 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
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
         CurrentDate     =   38434
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   315
         Left            =   5040
         TabIndex        =   5
         Top             =   300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
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
         CurrentDate     =   38434
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Inicio "
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
         Left            =   165
         TabIndex        =   12
         Top             =   345
         Width           =   510
      End
      Begin VB.Label Label4 
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
         Left            =   4140
         TabIndex        =   11
         Top             =   345
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
      Height          =   1110
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton cmdBuscarTrabajador 
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
         Height          =   255
         Left            =   6075
         TabIndex        =   3
         ToolTipText     =   "Búsqueda de Trabajador"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtApellido 
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
         TabIndex        =   2
         Top             =   690
         Width           =   4815
      End
      Begin VB.TextBox txtIdTrabajador 
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
         TabIndex        =   1
         Text            =   "0"
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Left            =   150
         TabIndex        =   7
         Top             =   720
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código "
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
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7140
      Top             =   1500
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
            Picture         =   "FrmRepConsolidadoxTrabajador.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepConsolidadoxTrabajador.frx":64EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CLBAcciones 
      Height          =   1890
      Left            =   6990
      TabIndex        =   8
      Top             =   90
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   3334
      BandCount       =   1
      Orientation     =   1
      _CBWidth        =   795
      _CBHeight       =   1890
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TLBAcciones"
      MinHeight1      =   735
      Width1          =   1200
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TLBAcciones 
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
Attribute VB_Name = "FrmRepConsolidadoxTrabajador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrformname As String = "Reporte de Consolidado x Trabajador"
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
End Sub

Private Sub Form_Load()
    Me.Caption = mstrformname
    Me.DtpFechaFin.Value = Date
    Me.DtpFechaInicio.Value = Date
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        If Me.TxtIdTrabajador.Text = "0" Then
            MsgBox "Es Obligatorio Seleccionar un Trabajador", vbInformation, mstrformname
        Else
            Imprimir
        End If
    Case 3
        Limpiar
End Select
End Sub
Public Sub Imprimir()
Dim objDatosTrab As Object
Dim rstDatosTrab As Object
Dim strMensaje As String

On Error GoTo ErrHandler
    If Me.TxtIdTrabajador.Text = 0 Then
        MsgBox "Debe seleccionar un trabajador", vbCritical, mstrformname
    Else
        Screen.MousePointer = vbHourglass
        FrmMensajeEspera.ShowForm
        FrmMensajeEspera.Refresh
        Set objDatosTrab = CreateObject("Personal.clsMovimientos")
        Set rstDatosTrab = objDatosTrab.GetConsolidadoxTrabajador(Val(Me.TxtIdTrabajador.Text), Me.DtpFechaInicio.Value, Me.DtpFechaFin.Value)
        Set objDatosTrab = Nothing
        If RecordsetReporteVacio(rstDatosTrab, strMensaje) Then
            MsgBox strMensaje, vbCritical, mstrformname
        Else
            gaParametrosReport(3, 1) = "FechaInicio"
            gaParametrosReport(3, 2) = CStr(DtpFechaInicio.Value)
            gaParametrosReport(4, 1) = "FechaFin"
            gaParametrosReport(4, 2) = CStr(DtpFechaFin.Value)
            ShowMultiReport rstDatosTrab, "rptRepConsolidadoTrabajador", gaParametrosReport()
        End If
        FrmMensajeEspera.Hide
        Set rstDatosTrab = Nothing
        Screen.MousePointer = vbDefault
    End If
    
Exit Sub
ErrHandler:
    FrmMensajeEspera.Hide
    Set rstDatosTrab = Nothing
    Set objDatosTrab = Nothing
    Screen.MousePointer = vbDefault
  
Call ErrorMessage(FrmRepConsolidadoxTrabajador_Imprimir, Err.Source & " FrmRepConsolidadoxTrabajador:Imprimir", Err.Description)
End Sub

Public Sub Limpiar()
    Me.TxtIdTrabajador.Text = 0
    Me.TxtApellido.Text = ""
    Me.DtpFechaInicio.Value = Date
    Me.DtpFechaFin.Value = Date
End Sub
