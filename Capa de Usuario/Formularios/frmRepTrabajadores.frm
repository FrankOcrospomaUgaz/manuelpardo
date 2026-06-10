VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmRepTrabajadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Trabajadores"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "frmRepTrabajadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6195
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
      Height          =   3090
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5130
      Begin VB.Frame Frame1 
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   195
         TabIndex        =   14
         Top             =   2220
         Width           =   4830
         Begin VB.OptionButton optNormal 
            Caption         =   "Normal"
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
            TabIndex        =   5
            Top             =   375
            Width           =   1020
         End
         Begin VB.OptionButton optSuspendido 
            Caption         =   "Suspendido"
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
            Left            =   1680
            TabIndex        =   6
            Top             =   375
            Width           =   1425
         End
         Begin VB.OptionButton optEliminado 
            Caption         =   "Eliminado"
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
            Left            =   3345
            TabIndex        =   7
            Top             =   375
            Width           =   1260
         End
      End
      Begin MSDataListLib.DataCombo dtcGrupo 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   3675
         _ExtentX        =   6482
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
      Begin MSDataListLib.DataCombo dtcCategoria 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   3675
         _ExtentX        =   6482
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
      Begin MSDataListLib.DataCombo dtcCargo 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   1215
         Width           =   3675
         _ExtentX        =   6482
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
      Begin MSDataListLib.DataCombo dtcArea 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   1680
         Width           =   3675
         _ExtentX        =   6482
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
         Left            =   195
         TabIndex        =   11
         Top             =   345
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Categoría"
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
         Left            =   195
         TabIndex        =   10
         Top             =   780
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
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
         Left            =   195
         TabIndex        =   9
         Top             =   1275
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Area"
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
         Left            =   195
         TabIndex        =   8
         Top             =   1740
         Width           =   405
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5475
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
            Picture         =   "frmRepTrabajadores.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepTrabajadores.frx":64EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   1890
      Left            =   5340
      TabIndex        =   12
      Top             =   150
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
Attribute VB_Name = "frmRepTrabajadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrformname As String = "Reporte de Trabajadores"

Public Sub ShowForm()
    Me.Show
    CenterForm Me
    Me.Caption = mstrformname
End Sub

Private Sub Form_Activate()
    Limpiar
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
Dim objPersonal As Object
    Screen.MousePointer = vbHourglass
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set Me.DtcCategoria.RowSource = objPersonal.GetCategoria()
    Set Me.dtcGrupo.RowSource = objPersonal.GetGrupo()
    Set Me.DtcCargo.RowSource = objPersonal.getcargo()
    Set Me.DtcArea.RowSource = objPersonal.GetArea()
    Set objPersonal = Nothing
    
    Me.dtcGrupo.ListField = "Descripcion"
    Me.dtcGrupo.BoundColumn = "IdGrupo"
    
    Me.DtcCategoria.ListField = "Descripcion"
    Me.DtcCategoria.BoundColumn = "IdCategoria"
    
    Me.DtcCargo.ListField = "Descripcion"
    Me.DtcCargo.BoundColumn = "IdCargo"
    
    Me.DtcArea.ListField = "Descripcion"
    Me.DtcArea.BoundColumn = "IdArea"
    
    Screen.MousePointer = vbDefault

Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Screen.MousePointer = vbDefault
    Call ErrorMessage(FrmRepTrabajadores_Load, Err.Source & " FrmRepTrabajadores :Load", Err.Description)
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

On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    FrmMensajeEspera.ShowForm
    FrmMensajeEspera.Refresh
    Set objReport = CreateObject("Personal.clsReportes")
    If Me.optNormal Then
        Set rstDataReport = objReport.GetReportTrabajadores(, gstrEstadoTrabajadorNormal, Val(Me.dtcGrupo.BoundText), _
        Val(Me.DtcCategoria.BoundText), Val(Me.DtcCargo.BoundText), Val(Me.DtcArea.BoundText))
        
        ElseIf Me.optSuspendido Then
            Set rstDataReport = objReport.GetReportTrabajadores(, gstrEstadoTrabajadorSuspendido, Val(Me.dtcGrupo.BoundText), _
            Val(Me.DtcCategoria.BoundText), Val(Me.DtcCargo.BoundText), Val(Me.DtcArea.BoundText))
            
            ElseIf Me.optEliminado Then
                Set rstDataReport = objReport.GetReportTrabajadores(, gstrEstadoTrabajadorEliminado, Val(Me.dtcGrupo.BoundText), _
                Val(Me.DtcCategoria.BoundText), Val(Me.DtcCargo.BoundText), Val(Me.DtcArea.BoundText))
    Else
        Set rstDataReport = objReport.GetReportTrabajadores(, , Val(Me.dtcGrupo.BoundText), _
        Val(Me.DtcCategoria.BoundText), Val(Me.DtcCargo.BoundText), Val(Me.DtcArea.BoundText))
    End If
    
    Set objReport = Nothing
    If RecordsetReporteVacio(rstDataReport, strMensaje) Then
        MsgBox strMensaje, vbCritical, Me.Caption
    Else
        ShowMultiReport rstDataReport, "rptRepListadoDeTrabajadores", gaParametrosReport()
    End If
    FrmMensajeEspera.Hide
    Set rstDataReport = Nothing
    Set objReport = Nothing
    Screen.MousePointer = vbDefault

Exit Sub
ErrHandler:
    FrmMensajeEspera.Hide
    Screen.MousePointer = vbDefault
    Set rstDataReport = Nothing
    Set objReport = Nothing
  
Call ErrorMessage(FrmRepTrabajadores_Imprimir, Err.Source & " FrmRepTrabajadores :Imprimir", Err.Description)

End Sub

Private Sub Limpiar()
    Me.DtcArea.BoundText = ""
    Me.DtcCargo.BoundText = ""
    Me.DtcCategoria.BoundText = ""
    Me.dtcGrupo.BoundText = ""
    Me.optEliminado.Value = False
    Me.optNormal.Value = False
    Me.optSuspendido.Value = False
End Sub
