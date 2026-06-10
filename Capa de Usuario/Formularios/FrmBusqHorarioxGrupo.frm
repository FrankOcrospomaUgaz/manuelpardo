VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBusqHorarioxGrupo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   Icon            =   "FrmBusqHorarioxGrupo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6585
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
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   5535
      Begin MSComCtl2.DTPicker DTPFechaInicio 
         Height          =   345
         Left            =   1170
         TabIndex        =   2
         Top             =   330
         Width           =   1560
         _ExtentX        =   2752
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
         CurrentDate     =   38415
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   345
         Left            =   3855
         TabIndex        =   3
         Top             =   330
         Width           =   1560
         _ExtentX        =   2752
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
         CurrentDate     =   38415
      End
      Begin VB.Label Label2 
         Caption         =   "Fin :"
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
         Left            =   3015
         TabIndex        =   9
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Inicio :"
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
         TabIndex        =   8
         Top             =   390
         Width           =   615
      End
   End
   Begin VB.Frame FraCriterio 
      Caption         =   "Criterio de Consulta"
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
      Left            =   75
      TabIndex        =   5
      Top             =   135
      Width           =   5535
      Begin MSDataListLib.DataCombo DtcGrupo 
         Height          =   315
         Left            =   1245
         TabIndex        =   1
         Top             =   345
         Width           =   4215
         _ExtentX        =   7435
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
      Begin VB.Label LblGrupo 
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
         TabIndex        =   6
         Top             =   390
         Width           =   510
      End
   End
   Begin MSComctlLib.ImageList ImgLstAcciones 
      Left            =   5040
      Top             =   -120
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
            Picture         =   "FrmBusqHorarioxGrupo.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBusqHorarioxGrupo.frx":35D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   1740
      Left            =   5655
      TabIndex        =   0
      Top             =   195
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   3069
      BandCount       =   1
      FixedOrder      =   -1  'True
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   855
      _CBHeight       =   1740
      _Version        =   "6.0.8169"
      BandBackColor1  =   -2147483635
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   795
      Width1          =   2955
      FixedBackground1=   0   'False
      UseCoolbarPicture1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   570
         Left            =   30
         TabIndex        =   4
         Top             =   345
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImgLstAcciones"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Consultar]"
               Object.ToolTipText     =   "Consultar datos"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Limpiar]"
               Object.ToolTipText     =   "Limpiar criterios de consulta"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmBusqHorarioxGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strModo As enumActions
Public intIdGrupo As Integer

Private mrstGrupo As Object

Public blnStatus As Boolean
Public strMensajeError As String

Private Const mstrformname As String = "Consulta de Horario de Grupo"

Public Property Get rstGrupo() As Variant
    Set rstGrupo = mrstGrupo
End Property
Public Property Set rstGrupo(ByVal vNewValue As Variant)
    Set mrstGrupo = vNewValue
    If Not rstGrupo.EOF Then
        Set Me.dtcGrupo.RowSource = rstGrupo
        Me.dtcGrupo.ListField = "Descripcion"
        Me.dtcGrupo.BoundColumn = "IdGrupo"
        Me.dtcGrupo.Enabled = True
    Else
        Me.dtcGrupo.Enabled = False
        Me.dtcGrupo.Text = ""
    End If
End Property

Public Sub ShowForm()
On Error GoTo ErrHandler
    LoadData
    If Me.strModo = Find Then
        Me.Show vbModal
    End If
    If Me.strModo = Query Then
        Me.Show
    End If
    Me.ZOrder

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmBusqHorarioxGrupo_ShowForm, Err.Source & " FrmBusqHorarioxGrupo:ShowForm", Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrstGrupo = Nothing
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case KEY_ACTION_CLEAR
            Limpiar
        Case KEY_ACTION_QUERY
            Queries
    End Select
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.strMensajeError = ""
    Me.Caption = mstrformname
End Sub

Private Sub LoadData()
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set rstGrupo = objPersonal.GetGrupo()
    Set objPersonal = Nothing
   
    Me.dtcGrupo.BoundText = Me.intIdGrupo
    Me.DtpFechaInicio.Value = Date
    Me.DtpFechaFin.Value = Date
    Me.dtcGrupo.SetFocus
    
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Call ErrorMessage(FrmBusqHorarioxGrupo_LoadData, Err.Source & " FrmBusqHorarioxGrupo:LoadData", Err.Description)
End Sub

Public Function ValidateFields() As Boolean
    ValidateFields = False
   
    If Me.dtcGrupo.BoundText = "" Then
        intIdGrupo = 0
    Else
        intIdGrupo = Me.dtcGrupo.BoundText
    End If
    
    ValidateFields = True
End Function

Private Function Queries() As Boolean

On Error GoTo ErrHandler:
    If Not ValidateFields Then
        Exit Function
    Else
        Screen.MousePointer = vbHourglass
        Queries = True
        FrmConsHorarioxGrupo.intIdGrupo = Me.intIdGrupo
        FrmConsHorarioxGrupo.datFechaInicio = Me.DtpFechaInicio.Value
        FrmConsHorarioxGrupo.datFechaFin = Me.DtpFechaFin.Value
        FrmConsHorarioxGrupo.strModo = Me.strModo
        FrmConsHorarioxGrupo.ShowForm
        Unload Me
        Screen.MousePointer = vbDefault
    End If

Exit Function
ErrHandler:
   Queries = False
   Screen.MousePointer = vbDefault

Call ErrorMessage(FrmBusqHorarioxGrupo_Queries, Err.Source & " FrmBusqHorarioxGrupo:Queries", Err.Description)
End Function

Private Function Limpiar() As Boolean
    Me.dtcGrupo.BoundText = ""
    Me.DtpFechaFin.Value = Date
    Me.DtpFechaInicio.Value = Date
    Me.dtcGrupo.SetFocus
    Limpiar = True
End Function





