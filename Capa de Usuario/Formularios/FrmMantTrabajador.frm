VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmMantTrabajador 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "FrmMantTrabajador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10230
   Begin VB.Frame FraRuta 
      Caption         =   "Ruta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   90
      TabIndex        =   33
      Top             =   5520
      Width           =   5835
      Begin VB.Label LblFoto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   90
         TabIndex        =   34
         Top             =   240
         Width           =   5610
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame FraAcciones 
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   6000
      TabIndex        =   31
      Top             =   5520
      Width           =   3360
      Begin VB.CommandButton cmdAceptar 
         Height          =   615
         Left            =   1020
         Picture         =   "FrmMantTrabajador.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   135
         Width           =   720
      End
      Begin VB.CommandButton cmdCancelar 
         Height          =   615
         Left            =   1770
         Picture         =   "FrmMantTrabajador.frx":190C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   150
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CmgRuta 
      Left            =   9720
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FraFoto 
      Caption         =   "Foto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5490
      Left            =   6000
      TabIndex        =   24
      Top             =   15
      Width           =   4200
      Begin VB.CheckBox chkIndTarjeta 
         Caption         =   "Permitir Tarjeta"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   5160
         Width           =   1935
      End
      Begin VB.CommandButton cmdProbarHuella 
         Caption         =   "&Probar Huella"
         Height          =   255
         Left            =   2160
         TabIndex        =   37
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdHuellaDigital 
         Caption         =   "Cargar Huella"
         Height          =   285
         Left            =   2160
         TabIndex        =   35
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton cmdDefinirRutaFoto 
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
         Height          =   285
         Left            =   3600
         TabIndex        =   15
         ToolTipText     =   "Definir archivo de foto"
         Top             =   4920
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Sin Huella"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Image ImgFoto 
         Height          =   4440
         Left            =   600
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3105
      End
   End
   Begin VB.Frame FraSecundario 
      Caption         =   "Datos de Trabajo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      Left            =   60
      TabIndex        =   21
      Top             =   2985
      Width           =   5880
      Begin VB.TextBox TxtTarjeta 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   9
         Tag             =   "2"
         Top             =   630
         Width           =   2130
      End
      Begin MSDataListLib.DataCombo DtcCategoria 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Tag             =   "2"
         Top             =   1335
         Width           =   4590
         _ExtentX        =   8096
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
      Begin MSDataListLib.DataCombo DtcCargo 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Tag             =   "2"
         Top             =   1725
         Width           =   4590
         _ExtentX        =   8096
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
      Begin MSDataListLib.DataCombo DtcArea 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Tag             =   "2"
         Top             =   2070
         Width           =   4590
         _ExtentX        =   8096
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
      Begin MSComCtl2.DTPicker DtpFechaIngreso 
         Height          =   330
         Left            =   1200
         TabIndex        =   8
         Tag             =   "2"
         Top             =   270
         Width           =   2115
         _ExtentX        =   3731
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
         Format          =   50462721
         CurrentDate     =   38036
      End
      Begin MSDataListLib.DataCombo DtcGrupo 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Tag             =   "2"
         Top             =   975
         Width           =   4590
         _ExtentX        =   8096
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
      Begin VB.Label LblIdGrupo 
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
         Left            =   135
         TabIndex        =   32
         Top             =   1005
         Width           =   510
      End
      Begin VB.Label LblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta"
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
         Left            =   135
         TabIndex        =   30
         Top             =   645
         Width           =   630
      End
      Begin VB.Label LblFechaIngreso 
         AutoSize        =   -1  'True
         Caption         =   "Ingreso"
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
         Left            =   135
         TabIndex        =   29
         Top             =   315
         Width           =   660
      End
      Begin VB.Label LblCategoria 
         AutoSize        =   -1  'True
         Caption         =   "Categoria"
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
         Left            =   135
         TabIndex        =   28
         Top             =   1335
         Width           =   825
      End
      Begin VB.Label LblCargo 
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
         Left            =   135
         TabIndex        =   27
         Top             =   1725
         Width           =   495
      End
      Begin VB.Label LblArea 
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
         Left            =   135
         TabIndex        =   26
         Top             =   2100
         Width           =   405
      End
   End
   Begin VB.Frame FraPrincipal 
      Caption         =   "Datos Personales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6105
      Begin VB.TextBox TxtNombre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   3
         Tag             =   "2"
         Top             =   960
         Width           =   4575
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
         Height          =   375
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   2
         Tag             =   "2"
         Top             =   600
         Width           =   4575
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
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Tag             =   "-1"
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtTelefono 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   6
         Tag             =   "2"
         Top             =   2040
         Width           =   4575
      End
      Begin VB.TextBox TxtDireccion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   5
         Tag             =   "2"
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox TxtDni 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   4
         Tag             =   "2"
         Top             =   1320
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DtpFechaNacimiento 
         Height          =   330
         Left            =   1155
         TabIndex        =   7
         Tag             =   "2"
         Top             =   2505
         Width           =   2145
         _ExtentX        =   3784
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
         Format          =   50462721
         CurrentDate     =   38036
      End
      Begin VB.Label LblFechaNacimiento 
         AutoSize        =   -1  'True
         Caption         =   "Nacimiento"
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
         Left            =   135
         TabIndex        =   25
         Top             =   2550
         Width           =   945
      End
      Begin VB.Label LblTelefono 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono"
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
         Left            =   135
         TabIndex        =   23
         Top             =   2175
         Width           =   735
      End
      Begin VB.Label LblDireccion 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
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
         Left            =   135
         TabIndex        =   22
         Top             =   1785
         Width           =   780
      End
      Begin VB.Label LblApellido 
         AutoSize        =   -1  'True
         Caption         =   "Apellido"
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
         Left            =   135
         TabIndex        =   20
         Top             =   615
         Width           =   675
      End
      Begin VB.Label LblNombre 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
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
         Left            =   135
         TabIndex        =   19
         Top             =   1005
         Width           =   660
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
         Left            =   135
         TabIndex        =   18
         Top             =   255
         Width           =   570
      End
      Begin VB.Label LblDni 
         AutoSize        =   -1  'True
         Caption         =   "DNI"
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
         Left            =   135
         TabIndex        =   17
         Top             =   1395
         Width           =   300
      End
   End
End
Attribute VB_Name = "FrmMantTrabajador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strAction As String
Public intIdTrabajador As Integer
Public strNombre As String
Public strApellido As String
Public strDni As String
Public strDireccion As String
Public strTelefono As String
Public strTarjeta As String
Public strFoto As String
Public strEstado As String
Public datFechaIngreso As Date
Public datFechaNacimiento As Date
Public intIdCategoria As Integer
Public intIdCargo As Integer
Public intIdArea As Integer
Public intIdGrupo As Integer

Private mrstCategoria As Object
Private mrstCargo As Object
Private mrstArea As Object
Private mrstGrupo As Object

Public blnStatus As Boolean
Public strMensajeError As String

Private Const mstrformname As String = "Mantenimiento de Trabajadores"

Public Property Get rstGrupo() As Variant
    Set rstGrupo = mrstGrupo
End Property
Public Property Set rstGrupo(ByVal vNewValue As Variant)
    Set mrstGrupo = vNewValue
    If Not rstGrupo.EOF Then
        Set Me.DtcGrupo.RowSource = rstGrupo
        Me.DtcGrupo.ListField = "Descripcion"
        Me.DtcGrupo.BoundColumn = "IdGrupo"
        Me.DtcGrupo.Enabled = True
    Else
        Me.DtcGrupo.Enabled = False
        Me.DtcGrupo.Text = ""
    End If
End Property

Public Property Get rstCategoria() As Variant
    Set rstCategoria = mrstCategoria
End Property
Public Property Set rstCategoria(ByVal vNewValue As Variant)
    Set mrstCategoria = vNewValue
    If Not rstCategoria.EOF Then
        Set Me.DtcCategoria.RowSource = rstCategoria
        Me.DtcCategoria.ListField = "Descripcion"
        Me.DtcCategoria.BoundColumn = "IdCategoria"
        Me.DtcCategoria.Enabled = True
    Else
        Me.DtcCategoria.Enabled = False
        Me.DtcCategoria.Text = ""
    End If
End Property

Public Property Get rstCargo() As Variant
    Set rstCargo = mrstCargo
End Property
Public Property Set rstCargo(ByVal vNewValue As Variant)
    Set mrstCargo = vNewValue
    If Not rstCargo.EOF Then
        Set Me.DtcCargo.RowSource = rstCargo
        Me.DtcCargo.ListField = "Descripcion"
        Me.DtcCargo.BoundColumn = "IdCargo"
        Me.DtcCargo.Enabled = True
    Else
        Me.DtcCargo.Enabled = False
        Me.DtcCargo.Text = ""
    End If
End Property

Public Property Get rstArea() As Variant
    Set rstArea = mrstArea
End Property
Public Property Set rstArea(ByVal vNewValue As Variant)
    Set mrstArea = vNewValue
    If Not rstArea.EOF Then
        Set Me.DtcArea.RowSource = rstArea
        Me.DtcArea.ListField = "Descripcion"
        Me.DtcArea.BoundColumn = "IdArea"
        Me.DtcArea.Enabled = True
    Else
        Me.DtcArea.Enabled = False
        Me.DtcArea.Text = ""
    End If
End Property

Public Sub ShowForm()
On Error GoTo ErrHandler
    Select Case strAction
        Case enumActions.Update
            LoadData
            SettingControls enumActions.Update, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_UPDATE
        Case enumActions.AddNew
            LoadData
            SettingControls enumActions.AddNew, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_ADDNEW
            Me.TxtIdTrabajador.Text = 0
            Me.TxtApellido.Text = ""
            Me.TxtNombre.Text = ""
  '          Me.TxtDni.Text = ""
            Me.TxtDireccion.Text = ""
 '   Me.TxtTelefono.Text = ""
            Me.DtpFechaIngreso.Value = Date
            Me.DtpFechaNacimiento.Value = Date
            Me.TxtTarjeta.Text = ""
            Me.DtcCategoria.Text = "Seminario"
            Me.DtcCargo.Text = "Asistente"
            Me.DtcArea.Text = "Seminario"
            Me.DtcGrupo.Text = "Seminario"
    End Select
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    
    Call ErrorMessage(FrmMantTrabajador_ShowForm, Err.Source & " FrmMantTrabajador:ShowForm", Err.Description)
End Sub

Private Sub CmdAceptar_Click()
Dim blnResult As Boolean
On Error GoTo ErrHandler
    If Not ValidateFields Then Exit Sub
    Select Case strAction
        Case enumActions.Update
            blnResult = Update()
        Case enumActions.AddNew
            blnResult = AddNew()
    End Select

    If Not blnResult Then
        MsgBox MSG_OPERATION_FAIL & vbCrLf & Me.strMensajeError, vbCritical, mstrformname
        Me.blnStatus = False
    Else
        FRegTemplate = ""
        Unload Me
        If ActiveForm("FrmConsTrabajador") Then
            FrmConsTrabajador.LoadData
        End If
        FrmConsTrabajador.Show
    End If

Exit Sub
ErrHandler:
Call ErrorMessage(FrmMantTrabajador_ShowForm, Err.Source & " FrmMantTrabajador:ShowForm", Err.Description)
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
    If ActiveForm("FrmConsTrabajador") Then
            FrmConsTrabajador.LoadData
    End If
    FrmConsTrabajador.Show
End Sub

Private Sub cmdDefinirRutaFoto_Click()
    Set Me.ImgFoto.Picture = Nothing
    CmgRuta.DialogTitle = "Busqueda de fotos"
    CmgRuta.InitDir = App.Path
    CmgRuta.Filter = "Todos los Archivos de Imagen|*.jpg;*.bmp;*.gif;*.jpeg|Archivos JPEG(*.jpg,*.jpeg)|*.jpg;*.jpeg|Archivos GIF(*.gif)|*.gif|Archivos BMP(*.bmp)|*.bmp|"
    CmgRuta.FilterIndex = 1
    CmgRuta.ShowOpen
    Set Me.ImgFoto.Picture = LoadPicture(CmgRuta.FileName)
    Me.LblFoto.Caption = CmgRuta.FileName
    Me.ImgFoto.Refresh
End Sub

Private Sub cmdHuellaDigital_Click()
    frmHuella.Show 1
    If FRegTemplate <> "" Then
        Me.Label1.Visible = False
    Else
        Me.Label1.Visible = True
    End If
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.strMensajeError = ""
    Me.Caption = mstrformname
    FRegTemplate = ""
End Sub

Private Sub LoadData()
Dim objPersonal As Object
Dim rstPersonal As Object
Dim mivar As Variant

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set rstPersonal = objPersonal.GetUnicoTrabajador(Me.intIdTrabajador)
    Set rstCategoria = objPersonal.GetCategoria()
    Set rstCargo = objPersonal.getcargo()
    Set rstArea = objPersonal.GetArea()
    Set rstGrupo = objPersonal.GetGrupo()
    Set objPersonal = Nothing
    
    If Not rstPersonal.EOF Then
        Me.DtcCategoria.BoundText = rstPersonal("IdCategoria")
        
        Me.DtcCargo.BoundText = rstPersonal("IdCargo")
        Me.DtcArea.BoundText = rstPersonal("IdArea")
        Me.DtcGrupo.BoundText = rstPersonal("IdGrupo")
        
        Me.TxtIdTrabajador.Text = rstPersonal("IdTrabajador")
        Me.TxtNombre.Text = rstPersonal("Nombre")
        Me.TxtApellido.Text = rstPersonal("Apellido")
        Me.TxtDni.Text = rstPersonal("Dni")
        Me.TxtDireccion.Text = rstPersonal("Direccion")
        Me.TxtTelefono.Text = rstPersonal("Telefono")
        Me.DtpFechaNacimiento.Value = rstPersonal("FechaNacimiento")
        Me.DtpFechaIngreso.Value = rstPersonal("FechaIngreso")
        Me.TxtTarjeta.Text = rstPersonal("Tarjeta")
        
         mivar = IsNull(rstPersonal("Huella"))
        If mivar = True Then
            Me.Label1.Caption = "Huella no Registrada"
            cmdProbarHuella.Enabled = False
        Else
            Me.Label1.Caption = "Huella Registrada"
            cmdProbarHuella.Enabled = True
            
        End If
        
        If rstPersonal("IndTarjeta") = "S" Then
            Me.chkIndTarjeta.Value = 1
        Else
            Me.chkIndTarjeta.Value = 0
        End If
        
        'If rstPersonal("Huella") <> "" Then
        '   Me.Label1.Visible = False
        'Else
        ' Me.Label1.Visible = True
        'End If
        
        If rstPersonal("Foto") <> "" Then
            'validar si la ruta existe
            'si existe
            Me.LblFoto.Caption = rstPersonal("Foto")
            Set Me.ImgFoto.Picture = LoadPicture(rstPersonal("Foto"))
            'si no existe muestra error
        End If
    End If
    Set rstPersonal = Nothing
    
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Set rstPersonal = Nothing
    Call ErrorMessage(FrmMantTrabajador_LoadData, Err.Source & " FrmMantTrabajador:LoadData", Err.Description)
End Sub

Private Function AddNew() As Boolean
Dim objPersonal As Object
Dim objGeneral As Object
Dim strMensajeError As String
Dim blnResult As Boolean

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    blnResult = objPersonal.AddNewTrabajador(Me.strTarjeta, Me.strNombre, Me.strApellido, _
                                    Me.strDni, Me.intIdCategoria, Me.intIdCargo, _
                                    Me.intIdArea, Me.intIdGrupo, "\\92.0.0.231\Personal\Insignia\In.jpg", Me.datFechaIngreso, _
                                    Me.datFechaNacimiento, Me.strTelefono, Me.strDireccion, FRegTemplate, IIf(Me.chkIndTarjeta.Value = 1, "S", "N"))
    Set objPersonal = Nothing
    If Not blnResult Then
        MsgBox strMensajeError, vbCritical, mstrformname
        AddNew = False
    Else
        AddNew = True
        Set objGeneral = CreateObject("Personal.clsGeneral")
        objGeneral.AddNewBitacora gstrNombreUsuario, "Trabajador ", KEY_ACTION_ADDNEW, ""
        Set objGeneral = Nothing
    End If

Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Set objGeneral = Nothing
    AddNew = False

Call ErrorMessage(FrmMantTrabajador_AddNew, Err.Source & " FrmMantTrabajador:AddNew", Err.Description)
End Function

Private Function Update() As Boolean
Dim objPersonal As Object
Dim objGeneral As Object
Dim strMensajeError As String
Dim blnResult As Boolean

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    blnResult = objPersonal.UpdateTrabajador(Me.intIdTrabajador, Me.strTarjeta, Me.strNombre, Me.strApellido, _
                                    Me.strDni, Me.intIdCategoria, Me.intIdCargo, _
                                    Me.intIdArea, Me.strEstado, Me.intIdGrupo, Me.strFoto, Me.datFechaIngreso, _
                                    Me.datFechaNacimiento, Me.strTelefono, Me.strDireccion, FRegTemplate, IIf(Me.chkIndTarjeta.Value = 1, "S", "N"))
    Set objPersonal = Nothing
    If Not blnResult Then
        MsgBox strMensajeError, vbCritical, mstrformname
        Update = False
    Else
        Update = True
        Set objGeneral = CreateObject("Personal.clsGeneral")
        objGeneral.AddNewBitacora gstrNombreUsuario, "Trabajador ", KEY_ACTION_UPDATE, ""
        Set objGeneral = Nothing
    End If

Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Set objGeneral = Nothing
    Update = False

Call ErrorMessage(FrmMantTrabajador_Update, Err.Source & " FrmMantTrabajador:Update", Err.Description)
End Function

Public Function ValidateFields() As Boolean
    ValidateFields = False
    If Trim$(TxtApellido.Text) = "" Then
        MsgBox "Debe ingresar el apellido", vbInformation, mstrformname
        TxtApellido.SetFocus
        Exit Function
    End If
    If Trim$(TxtDni.Text) = "" Then
        MsgBox "Debe ingresar el DNI", vbInformation, mstrformname
        TxtApellido.SetFocus
        Exit Function
    End If
    If Trim$(Me.DtcCategoria.Text) = "" Then
        MsgBox "Debe seleccionar una categoría", vbInformation, mstrformname
        DtcCategoria.SetFocus
        Exit Function
    End If
    If Trim$(Me.DtcCargo.Text) = "" Then
        MsgBox "Debe seleccionar un cargo", vbInformation, mstrformname
        DtcCargo.SetFocus
        Exit Function
    End If
    If Trim$(Me.DtcArea.Text) = "" Then
        MsgBox "Debe seleccionar un área ", vbInformation, mstrformname
        DtcArea.SetFocus
        Exit Function
    End If
    
    Me.intIdTrabajador = Val(Me.TxtIdTrabajador.Text)
    Me.strNombre = Replace(Me.TxtNombre.Text, "'", "''")
    Me.strApellido = Replace(Me.TxtApellido.Text, "'", "''")
    Me.strDni = Replace(Me.TxtDni.Text, "'", "''")
    Me.strDireccion = Replace(Me.TxtDireccion.Text, "'", "''")
    Me.strTelefono = Replace(Me.TxtTelefono.Text, "'", "''")
    Me.strTarjeta = Replace(Me.TxtTarjeta.Text, "'", "''")
    Me.datFechaIngreso = Me.DtpFechaIngreso.Value
    Me.datFechaNacimiento = Me.DtpFechaNacimiento.Value
    Me.intIdCategoria = Me.DtcCategoria.BoundText
    Me.intIdCargo = Me.DtcCargo.BoundText
    Me.intIdArea = Me.DtcArea.BoundText
    Me.intIdGrupo = Me.DtcGrupo.BoundText
    Me.strEstado = gstrEstadoTrabajadorNormal
    Me.strFoto = Me.LblFoto.Caption
    
    ValidateFields = True
End Function

Private Sub Form_Terminate()
    Set mrstCategoria = Nothing
    Set mrstCargo = Nothing
    Set mrstArea = Nothing
    Set mrstGrupo = Nothing
End Sub
