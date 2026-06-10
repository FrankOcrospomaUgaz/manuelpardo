VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form FrmRepConsolidado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Consolidados Mensuales"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   Icon            =   "FrmRepConsolidado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5355
   Begin VB.Frame FraMes 
      Caption         =   "MES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   90
      TabIndex        =   3
      Top             =   75
      Width           =   4215
      Begin VB.ComboBox cboMes 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmRepConsolidado.frx":08CA
         Left            =   195
         List            =   "FrmRepConsolidado.frx":08F2
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   3825
      End
   End
   Begin VB.Frame FraAño 
      Caption         =   "AÑO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   90
      TabIndex        =   0
      Top             =   1275
      Width           =   4215
      Begin VB.TextBox txtAño 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   1
         Top             =   390
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ejemplo: 1999, 2004"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   2
         Top             =   480
         Width           =   1485
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4770
      Top             =   1590
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
            Picture         =   "FrmRepConsolidado.frx":095A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepConsolidado.frx":657E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   1890
      Left            =   4440
      TabIndex        =   5
      Top             =   165
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
         TabIndex        =   6
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
Attribute VB_Name = "FrmRepConsolidado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrformname As String = "Reporte de Consolidados Mensuales"

Public Sub ShowForm()
    Me.Show
    CenterForm Me
    Me.Caption = mstrformname
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrHandler
    If Button.Index = 1 Then
        If Me.cboMes.Text <> "" And IsNumeric(Me.txtAño.Text) Then
            If CInt(Me.txtAño.Text) >= 1000 And CInt(Me.txtAño.Text) <= 9999 Then
                Imprimir
            Else
                MsgBox "Año Incorrecto", vbExclamation, "Mensaje"
            End If
        Else
            MsgBox "Debe Especificar un Mes y un Año Obligatorio", vbInformation, "Mensaje"
        End If
    Else
        Limpiar
    End If
    Exit Sub

ErrHandler:

    Call ErrorMessage(FrmRepConsolidado_Toolbar, Err.Source & " FrmRepConsolidado: Toolbar", Err.Description)
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
    Set rstDataReport = objReporte.GetReportConsolidados(Me.cboMes.Text, Me.txtAño.Text)
    Set objReporte = Nothing
    If RecordsetReporteVacio(rstDataReport, strMensaje) Then
        MsgBox strMensaje, vbCritical, mstrformname
    Else
        ShowMultiReport rstDataReport, "rptRepConsolidadoMensual", gaParametrosReport()
    End If
    FrmMensajeEspera.Hide
    Screen.MousePointer = vbDefault
    Set rstDataReport = Nothing
    
Exit Sub
ErrHandler:
    FrmMensajeEspera.Hide
    Set rstDataReport = Nothing
    Set objReporte = Nothing
    Screen.MousePointer = vbDefault
    
    Call ErrorMessage(FrmRepConsolidado_Imprimir, Err.Source & " FrmRepConsolidado:Imprimir", Err.Description)
End Sub
Private Sub Limpiar()
    Me.cboMes.ListIndex = -1
    Me.txtAño.Text = Empty
End Sub
