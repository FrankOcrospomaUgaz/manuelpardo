VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRepBitacora 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "FrmRepBitacora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6615
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
      Height          =   2625
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   5520
      Begin VB.ComboBox cboOperacion 
         Height          =   315
         ItemData        =   "FrmRepBitacora.frx":08CA
         Left            =   1275
         List            =   "FrmRepBitacora.frx":08D7
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1260
         Width           =   4050
      End
      Begin VB.ComboBox cboTabla 
         Height          =   315
         ItemData        =   "FrmRepBitacora.frx":08FD
         Left            =   1275
         List            =   "FrmRepBitacora.frx":092B
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   810
         Width           =   4050
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
         TabIndex        =   7
         Top             =   1770
         Width           =   5205
         Begin VB.CheckBox ChkFecha 
            Height          =   240
            Left            =   165
            TabIndex        =   4
            Top             =   -15
            Width           =   300
         End
         Begin MSComCtl2.DTPicker DtpFechaFin 
            Height          =   285
            Left            =   3480
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
            Format          =   63700993
            CurrentDate     =   38056
         End
         Begin MSComCtl2.DTPicker DtpFechaInicio 
            Height          =   285
            Left            =   1155
            TabIndex        =   5
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
            Format          =   63700993
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
            Left            =   3015
            TabIndex        =   9
            Top             =   300
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
            TabIndex        =   8
            Top             =   315
            Width           =   465
         End
      End
      Begin MSDataListLib.DataCombo dtcUsuario 
         Height          =   315
         Left            =   1275
         TabIndex        =   1
         Top             =   330
         Width           =   4050
         _ExtentX        =   7144
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
         Caption         =   "Usuario :"
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
         Top             =   390
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Tabla :"
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
         Left            =   195
         TabIndex        =   11
         Top             =   795
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Operación :"
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
         Left            =   210
         TabIndex        =   10
         Top             =   1290
         Width           =   1050
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5805
      Top             =   2235
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
            Picture         =   "FrmRepBitacora.frx":09C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepBitacora.frx":65E7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   1890
      Left            =   5730
      TabIndex        =   12
      Top             =   240
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
Attribute VB_Name = "FrmRepBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrformname As String = "Reporte de Bitácora"
Public Sub ShowForm()
    Me.Show
    CenterForm Me
    Me.Caption = mstrformname
End Sub

Private Sub Form_Load()
    LoadData
End Sub

Public Sub LoadData()
Dim objUsuario As Object
On Error GoTo ErrHandler
    Set objUsuario = CreateObject("Personal.clsSeguridad")
    Set Me.dtcUsuario.RowSource = objUsuario.GetUsuario()
    Set objUsuario = Nothing
    Me.dtcUsuario.ListField = "NombreUsuario"
    Me.dtcUsuario.BoundColumn = "IdUsuario"
Exit Sub
ErrHandler:
    Set objUsuario = Nothing
    Call ErrorMessage(FrmRepBitacora_Load, Err.Source & " FrmRepBitacora : LoadData", Err.Description)
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
Dim rstDataReport As Object
Dim strMensaje As String
On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    FrmMensajeEspera.ShowForm
    FrmMensajeEspera.Refresh
    Set objReporte = CreateObject("Personal.clsReportes")
    If Me.ChkFecha.Value = 1 Then
        Set rstDataReport = objReporte.GetBitacora(Me.dtcUsuario.Text, Me.cboTabla.Text, Me.cboOperacion.Text, _
                                           Me.DtpFechaInicio.Value, Me.DtpFechaFin.Value)
    Else
        Set rstDataReport = objReporte.GetBitacora(Me.dtcUsuario.Text, Me.cboTabla.Text, Me.cboOperacion.Text)
    End If
    Set objReporte = Nothing
    If RecordsetReporteVacio(rstDataReport, strMensaje) Then
        MsgBox strMensaje, vbCritical, Me.Caption
    Else
        gaParametrosReport(3, 1) = "FechaInicio"
        gaParametrosReport(3, 2) = CStr(Me.DtpFechaInicio.Value)
        gaParametrosReport(4, 1) = "FechaFin"
        gaParametrosReport(4, 2) = CStr(Me.DtpFechaFin.Value)
        ShowMultiReport rstDataReport, "rptRepBitacora", gaParametrosReport()
    End If
    FrmMensajeEspera.Hide
    Set rstDataReport = Nothing
    Screen.MousePointer = vbDefault
Exit Sub
ErrHandler:
    FrmMensajeEspera.Hide
    Set rstDataReport = Nothing
    Set objReporte = Nothing
    Screen.MousePointer = vbDefault
    Call ErrorMessage(FrmRepBitacora_Imprimir, Err.Source & " FrmRepBitacora : Imprimir", Err.Description)
End Sub

Private Sub Limpiar()
    Me.dtcUsuario.Text = ""
    Me.cboOperacion.ListIndex = -1
    Me.cboTabla.ListIndex = -1
    Me.DtpFechaInicio.Value = Date
    Me.DtpFechaFin.Value = Date
    Me.ChkFecha.Value = 0
End Sub
