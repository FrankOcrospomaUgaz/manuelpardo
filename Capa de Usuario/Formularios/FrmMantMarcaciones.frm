VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMantMarcaciones 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   Icon            =   "FrmMantMarcaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7110
   Begin VB.Frame FraAcciones 
      Caption         =   "Acciones"
      Height          =   810
      Left            =   938
      TabIndex        =   9
      Top             =   2640
      Width           =   5235
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   2828
         Picture         =   "FrmMantMarcaciones.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancelar la operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   1680
         Picture         =   "FrmMantMarcaciones.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
   End
   Begin VB.Frame fraMarcacion 
      Height          =   2415
      Left            =   248
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtFecha 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   17
         Tag             =   "-1"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtIdMarcacion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Tag             =   "-1"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtIdTrabajador 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         TabIndex        =   13
         Tag             =   "-1"
         Top             =   360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpHoraIngresoIdeal 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Tag             =   "-1"
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "hh:mm tt"
         Format          =   63700995
         UpDown          =   -1  'True
         CurrentDate     =   38405
      End
      Begin MSComCtl2.DTPicker dtpHoraIngresoReal 
         Height          =   315
         Left            =   4920
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "hh:mm tt"
         Format          =   63700995
         UpDown          =   -1  'True
         CurrentDate     =   38405
      End
      Begin MSComCtl2.DTPicker dtpHoraSalidaIdeal 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Tag             =   "-1"
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "hh:mm tt"
         Format          =   63700995
         UpDown          =   -1  'True
         CurrentDate     =   38405
      End
      Begin MSComCtl2.DTPicker dtpHoraSalidaReal 
         Height          =   315
         Left            =   4920
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "hh:mm tt"
         Format          =   63700995
         UpDown          =   -1  'True
         CurrentDate     =   38405
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000016&
         X1              =   0
         X2              =   6600
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   16
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Id Marcación :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Id Trabajador :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Hora Salida Real :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Hora Salida Ideal :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Hora Ingreso Real :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Hora Ingreso Ideal :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   1440
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmMantMarcaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public intIdTrabajador As Integer
Public intIdMarcacion As Long
Public datFecha As Date
Public datHoraSalidaReal As Date
Public datHoraSalidaIdeal As Date
Public datHoraEntradaReal As Date
Public datHoraEntradaIdeal As Date

Public strMensajeError As String

Private Const mstrformname As String = "Mantenimiento de Marcaciones"
Public Sub ShowForm()
On Error GoTo ErrHandler
    Select Case intAction
        Case enumActions.Update
            LoadData
            SettingControls enumActions.Update, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_UPDATE
    End Select
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantCargo_ShowForm, Err.Source & " FrmMantCargo:ShowForm", Err.Description)
End Sub

Private Sub CmdAceptar_Click()
Dim blnResult As Boolean

On Error GoTo ErrHandler
    If Not ValidateFields Then Exit Sub
    Select Case intAction
        Case enumActions.Update
            blnResult = Update()
    End Select
    
    If Not blnResult Then
        MsgBox MSG_OPERATION_FAIL & vbCrLf & Me.strMensajeError, vbCritical, mstrformname
        'Me.blnStatus = False
    Else
        Unload Me
        If ActiveForm("FrmConsMarcaciones") Then
            FrmConsMarcaciones.LoadData
        End If
    End If

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantMarcaciones_CmdAceptar_Click, Err.Source & " FrmMantCargo:ShowForm", Err.Description)
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
    
        Me.TxtIdTrabajador.Text = Me.intIdTrabajador
        Me.txtIdMarcacion.Text = Me.intIdMarcacion
        Me.txtFecha.Text = Me.datFecha
        Me.DtpHoraIngresoIdeal.Value = Me.datHoraEntradaIdeal
        Me.DtpHoraIngresoReal.Value = Me.datHoraEntradaReal
        Me.DtpHoraSalidaIdeal.Value = Me.datHoraSalidaIdeal
        Me.DtpHoraSalidaReal.Value = Me.datHoraSalidaReal
    
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Set rstPersonal = Nothing
    Call ErrorMessage(FrmMantMarcaciones_LoadData, Err.Source & " FrmMantCargo:LoadData", Err.Description)
End Sub

Private Function Update() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsMovimientos")
    Update = objPersonal.UpdateHoraEntradaYSalidaMarcacion(Me.intIdMarcacion, Me.datHoraEntradaReal, Me.datHoraSalidaReal)
    Set objPersonal = Nothing
    If Update Then
        Bitacora "Marcacion", KEY_ACTION_UPDATE, "Codigo " & Str(Me.intIdMarcacion) & " - Modificacion de Marcacion "
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Update = False

Call ErrorMessage(FrmMantMarcaciones_Update, Err.Source & " FrmMantCargo:Update", Err.Description)
End Function

Public Function ValidateFields() As Boolean
    ValidateFields = False
    'If Trim$(TxtDescripcion.Text) = "" Then
    '    MsgBox MSG_REQUERY_FIELD, vbInformation, mstrFormName
    '    TxtDescripcion.SetFocus
    '    Exit Function
    'End If
    Me.intIdMarcacion = Val(Me.txtIdMarcacion.Text)
    Me.datHoraEntradaReal = Me.DtpHoraIngresoReal.Value
    Me.datHoraSalidaReal = Me.DtpHoraSalidaReal.Value
    ValidateFields = True
End Function




