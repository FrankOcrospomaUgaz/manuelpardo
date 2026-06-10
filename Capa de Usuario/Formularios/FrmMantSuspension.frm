VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMantSuspension 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   Icon            =   "FrmMantSuspension.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6090
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Suspensión"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   90
      TabIndex        =   13
      Top             =   1065
      Width           =   5910
      Begin VB.TextBox TxtIdSuspension 
         Height          =   285
         Left            =   1740
         TabIndex        =   4
         Tag             =   "-1"
         Top             =   300
         Width           =   1395
      End
      Begin VB.TextBox TxtMotivo 
         Height          =   870
         Left            =   1740
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1125
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker DtpFechaFin 
         Height          =   330
         Left            =   4365
         TabIndex        =   6
         Top             =   690
         Width           =   1425
         _ExtentX        =   2514
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
         Format          =   21364737
         CurrentDate     =   38061
      End
      Begin MSComCtl2.DTPicker DtpFechaInicio 
         Height          =   330
         Left            =   1740
         TabIndex        =   5
         Top             =   690
         Width           =   1425
         _ExtentX        =   2514
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
         Format          =   21364737
         CurrentDate     =   38061
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
         Left            =   3900
         TabIndex        =   17
         Top             =   720
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
         Left            =   165
         TabIndex        =   16
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Motivo"
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
         TabIndex        =   15
         Top             =   1170
         Width           =   585
      End
      Begin VB.Label LblIdSuspension 
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
         Left            =   165
         TabIndex        =   14
         Top             =   360
         Width           =   570
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
      Left            =   90
      TabIndex        =   12
      Top             =   3180
      Width           =   5925
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   2325
         Picture         =   "FrmMantSuspension.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   3240
         Picture         =   "FrmMantSuspension.frx":190C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancelar la operación"
         Top             =   210
         Width           =   810
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
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   5955
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
         Left            =   5325
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
         Tag             =   "-1"
         Top             =   615
         Width           =   4065
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
         Tag             =   "-1"
         Text            =   "0"
         Top             =   210
         Width           =   1215
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
         TabIndex        =   11
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
         Left            =   90
         TabIndex        =   10
         Top             =   270
         Width           =   570
      End
   End
End
Attribute VB_Name = "FrmMantSuspension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public intIdSuspension As Long
Public intIdTrabajador As Integer
Public datFechaInicio As Date
Public datFechaFin As Date
Public strMotivo As String
Public blnStatus As Boolean
Public strMensajeError As String

Private Const mstrformname As String = "Mantenimiento de Suspensión"

Public Sub ShowForm()
On Error GoTo ErrHandler
    Select Case intAction
        Case enumActions.Update
            LoadData
            SettingControls enumActions.Update, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_UPDATE
        Case enumActions.AddNew
            LoadData
            SettingControls enumActions.AddNew, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_ADDNEW
            Me.TxtIdSuspension.Text = 0
            Me.TxtIdTrabajador.Text = 0
            Me.TxtApellido.Text = ""
            Me.TxtMotivo.Text = ""
            Me.DtpFechaInicio.Value = Date
            Me.DtpFechaFin.Value = Date
    End Select
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantSuspension_ShowForm, Err.Source & " FrmMantSuspension:ShowForm", Err.Description)
End Sub

Private Sub CmdAceptar_Click()
Dim blnResult As Boolean

On Error GoTo ErrHandler
    If Not ValidateFields Then Exit Sub
    Select Case intAction
        Case enumActions.Update
            blnResult = Update()
        Case enumActions.AddNew
            blnResult = AddNew()
    End Select
    
    If Not blnResult Then
        MsgBox MSG_OPERATION_FAIL & vbCrLf & Me.strMensajeError, vbCritical, mstrformname
        Me.blnStatus = False
    Else
        Unload Me
        If ActiveForm("FrmConsSuspension") Then
            FrmConsSuspension.LoadData
        End If
    End If

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantSuspension_ShowForm, Err.Source & " FrmMantSuspension:ShowForm", Err.Description)
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

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.strMensajeError = ""
End Sub

Private Sub LoadData()
Dim objPersonal As Object
Dim rstPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsMovimientos")
    Set rstPersonal = objPersonal.GetSuspension(Me.intIdSuspension)
    Set objPersonal = Nothing
    If Not rstPersonal.EOF Then
        Me.TxtIdSuspension.Text = rstPersonal("IdSuspension")
        Me.TxtIdTrabajador.Text = rstPersonal("IdTrabajador")
        Me.TxtApellido.Text = rstPersonal("Apellido")
        Me.TxtMotivo.Text = rstPersonal("Motivo")
        Me.DtpFechaInicio.Value = rstPersonal("FechaInicio")
        Me.DtpFechaFin.Value = rstPersonal("FechaFin")
    End If
    Set rstPersonal = Nothing

Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Set rstPersonal = Nothing
    Call ErrorMessage(FrmMantSuspension_LoadData, Err.Source & " FrmMantSuspension:LoadData", Err.Description)
End Sub

Private Function AddNew() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsMovimientos")
    AddNew = objPersonal.AddNewSuspension(Me.intIdTrabajador, Me.datFechaInicio, _
                                        Me.datFechaFin, Me.strMotivo)
    Set objPersonal = Nothing
    If AddNew Then
        Bitacora "Suspension", KEY_ACTION_ADDNEW
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    AddNew = False

Call ErrorMessage(FrmMantSuspension_AddNew, Err.Source & " FrmMantSuspension:AddNew", Err.Description)
End Function
Private Function Update() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsMovimientos")
    Update = objPersonal.UpdateSuspension(Me.intIdSuspension, Me.datFechaInicio, _
                                        Me.datFechaFin, Me.strMotivo)
    Set objPersonal = Nothing
    If Update Then
        Bitacora "Suspension", KEY_ACTION_UPDATE, "Codigo " & Str(Me.intIdSuspension)
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Update = False
    Call ErrorMessage(FrmMantSuspension_Update, Err.Source & " FrmMantSuspension:Update", Err.Description)
End Function

Public Function ValidateFields() As Boolean
    ValidateFields = False
    If Me.TxtIdTrabajador.Text = 0 Then
        MsgBox MSG_REQUERY_FIELD, vbInformation, mstrformname
        Exit Function
    End If
    If Me.DtpFechaInicio.Value - Me.DtpFechaFin.Value > 0 Then
        MsgBox MSG_INVALIDATE_DATE, vbInformation, mstrformname
        DtpFechaInicio.SetFocus
        Exit Function
    End If
    If Not ValidateConsolidado Then
        MsgBox "La Fecha de Inicio de la Suspensión no es válida, es posible que se haya generado consolidado en ese mes", vbInformation, mstrformname
        Exit Function
    End If
    Me.intIdSuspension = Val(Me.TxtIdSuspension.Text)
    Me.intIdTrabajador = Val(Me.TxtIdTrabajador.Text)
    Me.strMotivo = Trim$(Me.TxtMotivo.Text)
    Me.datFechaInicio = Me.DtpFechaInicio.Value
    Me.datFechaFin = Me.DtpFechaFin.Value
    ValidateFields = True
End Function

Public Function ValidateConsolidado() As Boolean
Dim objConsolidado As Object
On Error GoTo ErrHandler
    Set objConsolidado = CreateObject("Personal.clsMovimientos")
    ValidateConsolidado = Not objConsolidado.GetExisteConsolidado(Me.DtpFechaInicio.Value)
    Set objConsolidado = Nothing

Exit Function
ErrHandler:
    Set objConsolidado = Nothing
    Call ErrorMessage(FrmMantSuspension_ValidateConsolidado, Err.Source & " FrmMantSuspension:ValidateConsolidado", Err.Description)
End Function
