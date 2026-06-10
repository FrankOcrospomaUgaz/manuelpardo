VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBusqEliminarHorario 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "FrmBusqEliminarHorario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7590
   Begin VB.Frame Frame1 
      Caption         =   "Criterio de Selección"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3045
      Left            =   105
      TabIndex        =   7
      Top             =   945
      Width           =   6345
      Begin VB.Frame Frame2 
         Caption         =   "Por Grupo"
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
         Left            =   180
         TabIndex        =   16
         Top             =   2040
         Width           =   6060
         Begin MSDataListLib.DataCombo dtcGrupo 
            Height          =   315
            Left            =   1770
            TabIndex        =   17
            Top             =   315
            Width           =   4200
            _ExtentX        =   7408
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
            Caption         =   "Grupo "
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
            Left            =   150
            TabIndex        =   18
            Top             =   330
            Width           =   615
         End
      End
      Begin VB.Frame FraTrabajador 
         Caption         =   "Por Trabajador"
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
         Left            =   180
         TabIndex        =   10
         Top             =   855
         Width           =   6060
         Begin VB.CommandButton CmdBuscarTrabajador 
            Caption         =   "..."
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
            Height          =   240
            Left            =   5460
            TabIndex        =   13
            ToolTipText     =   "Búsqueda de Trabajador"
            Top             =   255
            Width           =   465
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
            TabIndex        =   12
            Top             =   615
            Width           =   4215
         End
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
            TabIndex        =   11
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
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   270
            Width           =   570
         End
      End
      Begin VB.OptionButton OptGrupo 
         Caption         =   "Por Grupo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3270
         TabIndex        =   9
         Top             =   345
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton OptTrabajador 
         Caption         =   "Por Trabajador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   285
         TabIndex        =   8
         Top             =   345
         Width           =   1635
      End
   End
   Begin VB.Frame FraCriterios 
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
      Height          =   885
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   6285
      Begin MSComCtl2.DTPicker DTPFechaInicio 
         Height          =   330
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
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
         CurrentDate     =   38426
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   330
         Left            =   4485
         TabIndex        =   4
         Top             =   360
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
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
         CurrentDate     =   38426
      End
      Begin VB.Label Label3 
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
         Left            =   3975
         TabIndex        =   3
         Top             =   420
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Left            =   360
         TabIndex        =   1
         Top             =   420
         Width           =   465
      End
   End
   Begin MSComctlLib.ImageList ImgLstAcciones 
      Left            =   6720
      Top             =   2640
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
            Picture         =   "FrmBusqEliminarHorario.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBusqEliminarHorario.frx":35D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   2175
      Left            =   6600
      TabIndex        =   5
      Top             =   210
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   3836
      BandCount       =   1
      FixedOrder      =   -1  'True
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   855
      _CBHeight       =   2175
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
         TabIndex        =   6
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
Attribute VB_Name = "FrmBusqEliminarHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intIdGrupo As Integer
Public intIdTrabajador As Integer
Public datFechaInicio As Date
Public datFechaFin As Date

Private mrstGrupo As Object

Private Const mstrformname As String = "Búsqueda de Horarios A Eliminar"

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
    CenterForm Me
    Me.Caption = mstrformname
End Sub

Private Sub OptGrupo_Click()
    Me.CmdBuscarTrabajador.Enabled = False
    Me.TxtIdTrabajador.Text = 0
    Me.TxtApellido.Text = ""
    Me.dtcGrupo.BoundText = ""
    Me.dtcGrupo.Enabled = True
End Sub


Private Sub OptTrabajador_Click()
    Me.CmdBuscarTrabajador.Enabled = True
    Me.dtcGrupo.BoundText = ""
    Me.dtcGrupo.Enabled = False
End Sub

Public Sub ShowForm()
On Error GoTo ErrHandler
    LoadData
    Me.ZOrder

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmBusqHorarioxTrabajador_ShowForm, Err.Source & " FrmBusqHorarioxTrabajador:ShowForm", Err.Description)
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

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_ACTION_QUERY
        Queries
    Case KEY_ACTION_CLEAR
        Clear
End Select
End Sub
Public Sub Clear()
    Me.DtpFechaFin.Value = Date
    Me.DtpFechaInicio.Value = Date
    Me.TxtIdTrabajador.Text = 0
    Me.TxtApellido.Text = ""
    Me.dtcGrupo.BoundText = ""
End Sub
Public Function Queries() As Boolean
    If Not ValidateFields Then Exit Function
    FrmConsEliminarHorario.intIdGrupo = Me.intIdGrupo
    FrmConsEliminarHorario.intIdTrabajador = Me.intIdTrabajador
    FrmConsEliminarHorario.datFechaFin = Me.datFechaFin
    FrmConsEliminarHorario.datFechaInicio = Me.datFechaInicio
    FrmConsEliminarHorario.ShowForm
End Function
Public Function ValidateFields() As Boolean
    ValidateFields = False
    If Me.DtpFechaInicio.Value < Date Then
        MsgBox MSG_INVALIDATE_DATEFIRST, vbInformation, mstrformname
        Exit Function
    End If
    If Me.DtpFechaInicio.Value > Me.DtpFechaFin.Value Then
        MsgBox MSG_INVALIDATE_DATE, vbInformation, mstrformname
        Exit Function
    End If
    If Me.OptGrupo.Value Then
        If Me.dtcGrupo.BoundText = "" Then Exit Function
        Me.intIdGrupo = Val(Me.dtcGrupo.BoundText)
        Me.intIdTrabajador = 0
    Else
        If Me.TxtIdTrabajador.Text = 0 Then Exit Function
        Me.intIdTrabajador = Val(Me.TxtIdTrabajador.Text)
        Me.intIdGrupo = 0
    End If
       Me.datFechaFin = Me.DtpFechaFin.Value
       Me.datFechaInicio = Me.DtpFechaInicio.Value
    ValidateFields = True
End Function
