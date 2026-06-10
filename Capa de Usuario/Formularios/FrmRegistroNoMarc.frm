VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRegistroNoMarc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de No Marcaciones"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "FrmRegistroNoMarc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5595
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4335
      Begin MSDataListLib.DataCombo dtcGrupo 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label2 
         Caption         =   "Grupo:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fecha Inicio:"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4335
      Begin MSComCtl2.DTPicker DTPFin 
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   24379393
         CurrentDate     =   38381
      End
      Begin MSComCtl2.DTPicker DTPInicio 
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   24379393
         CurrentDate     =   38381
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Desde:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "FrmRegistroNoMarc.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroNoMarc.frx":64EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   1500
      Left            =   4560
      TabIndex        =   8
      Top             =   360
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   2646
      BandCount       =   1
      Orientation     =   1
      _CBWidth        =   960
      _CBHeight       =   1500
      _Version        =   "6.0.8169"
      MinHeight1      =   900
      Width1          =   1440
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   630
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1111
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Vista Previa"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Limpiar"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmRegistroNoMarc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim objPersonal As Object
Set objPersonal = CreateObject("Personal.clsPersonal")
Set Me.dtcGrupo.RowSource = objPersonal.getGrupo()
Me.dtcGrupo.ListField = "Descripcion"
Me.dtcGrupo.BoundColumn = "IdGrupo"
Me.DTPInicio.Value = Date
Me.DTPFin.Value = Date
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
    IMPRIMIR
Else
    LIMPIAR
End If
End Sub

Private Sub IMPRIMIR()
Dim objReport As Object
Dim rstDataReport As Object
Dim strMensaje As String

On Error GoTo Errhandler
    Screen.MousePointer = vbHourglass
    'AlmacenaCriterios
    'IMPRIMIR = True
    
    Set objReport = CreateObject("Personal.clsReportes")
    
    
    Set rstDataReport = objReport.GetRegistroNoMarcaciones(Me.dtcGrupo.BoundText, Me.DTPInicio.Value, Me.DTPFin.Value)
        Set objReport = Nothing
        If RecordsetReporteVacio(rstDataReport, strMensaje) Then
            MsgBox strMensaje, vbCritical, Me.Caption
        Else
            ShowMultiReport rstDataReport, "rptRepRegistroDeNoMarcaciones", gaParametrosReport()
        End If
    
    Set rstDataReport = Nothing
    Screen.MousePointer = vbDefault

Exit Sub
Errhandler:
    Set rstDataReport = Nothing
    Set objReport = Nothing
    Screen.MousePointer = vbDefault
    
Call ErrorMessage(FrmRegistroNoMarc_Imprimir, Err.Source & " FrmRegistroNoMarc:Imprimir", Err.Description)
End Sub

Private Sub LIMPIAR()
    Me.dtcGrupo.BoundText = ""
End Sub
