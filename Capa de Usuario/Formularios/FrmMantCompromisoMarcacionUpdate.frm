VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMantCompromisoMarcacionUpdate 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "FrmMantCompromisoMarcacionUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   5910
   Begin VB.Frame FraTrabajador 
      Caption         =   "Datos de Trabajador"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   30
      TabIndex        =   12
      Top             =   15
      Width           =   5835
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
         Left            =   1080
         TabIndex        =   0
         Tag             =   "-1"
         Text            =   "0"
         Top             =   315
         Width           =   1560
      End
      Begin VB.TextBox TxtNombres 
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
         Left            =   1080
         MaxLength       =   120
         TabIndex        =   1
         Tag             =   "-1"
         Top             =   675
         Width           =   4635
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
         Top             =   300
         Width           =   570
      End
      Begin VB.Label LblNombres 
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
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
         TabIndex        =   13
         Top             =   660
         Width           =   750
      End
   End
   Begin VB.Frame FraMarcacion 
      Caption         =   "Datos de Marcación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Left            =   45
      TabIndex        =   9
      Top             =   1065
      Width           =   5820
      Begin VB.Frame FraReal 
         Caption         =   "Horario registrado"
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
         Left            =   150
         TabIndex        =   18
         Top             =   2010
         Width           =   5550
         Begin VB.OptionButton OptAmbas 
            Caption         =   "Ambas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4530
            TabIndex        =   22
            Top             =   300
            Width           =   960
         End
         Begin VB.OptionButton OptNinguno 
            Caption         =   "Ninguno"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   135
            TabIndex        =   21
            Top             =   300
            Width           =   1110
         End
         Begin VB.OptionButton OptSoloIngreso 
            Caption         =   "Sólo Ingreso"
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
            Left            =   1275
            TabIndex        =   20
            Top             =   300
            Width           =   1530
         End
         Begin VB.OptionButton OptSoloSalida 
            Caption         =   "Sólo Salida"
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
            Left            =   2985
            TabIndex        =   19
            Top             =   300
            Width           =   1380
         End
         Begin MSComCtl2.DTPicker DtpHoraSalidaReal 
            Height          =   285
            Left            =   4035
            TabIndex        =   23
            Top             =   705
            Width           =   1335
            _ExtentX        =   2355
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
            CustomFormat    =   "hh:mm tt"
            Format          =   20971523
            UpDown          =   -1  'True
            CurrentDate     =   38056
         End
         Begin MSComCtl2.DTPicker DtpHoraIngresoReal 
            Height          =   285
            Left            =   960
            TabIndex        =   24
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
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
            CustomFormat    =   "hh:mm tt"
            Format          =   20971523
            UpDown          =   -1  'True
            CurrentDate     =   38056
         End
         Begin VB.Label LblHoraSalidaReal 
            AutoSize        =   -1  'True
            Caption         =   "Salida"
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
            Left            =   3420
            TabIndex        =   26
            Top             =   750
            Width           =   510
         End
         Begin VB.Label LblHoraIngresoReal 
            AutoSize        =   -1  'True
            Caption         =   " Ingreso"
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
            TabIndex        =   25
            Top             =   765
            Width           =   705
         End
      End
      Begin VB.Frame FraIdeal 
         Caption         =   "Horario por cumplir"
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
         Left            =   150
         TabIndex        =   15
         Top             =   1110
         Width           =   5520
         Begin MSComCtl2.DTPicker DtpHoraSalidaIdeal 
            Height          =   285
            Left            =   4035
            TabIndex        =   5
            Tag             =   "-1"
            Top             =   330
            Width           =   1335
            _ExtentX        =   2355
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
            CustomFormat    =   "hh:mm tt"
            Format          =   20971523
            UpDown          =   -1  'True
            CurrentDate     =   38056
         End
         Begin MSComCtl2.DTPicker DtpHoraIngresoIdeal 
            Height          =   285
            Left            =   960
            TabIndex        =   4
            Tag             =   "-1"
            Top             =   330
            Width           =   1335
            _ExtentX        =   2355
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
            CustomFormat    =   "hh:mm tt"
            Format          =   20971523
            UpDown          =   -1  'True
            CurrentDate     =   38056
         End
         Begin VB.Label LblHoraIngresoIdeal 
            AutoSize        =   -1  'True
            Caption         =   " Ingreso"
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
            TabIndex        =   17
            Top             =   375
            Width           =   705
         End
         Begin VB.Label LblHoraSalidaIdeal 
            AutoSize        =   -1  'True
            Caption         =   "Salida"
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
            Left            =   3390
            TabIndex        =   16
            Top             =   375
            Width           =   510
         End
      End
      Begin VB.TextBox TxtIdMarcacion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1065
         TabIndex        =   2
         Tag             =   "-1"
         Top             =   315
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   300
         Left            =   1065
         TabIndex        =   3
         Tag             =   "-1"
         Top             =   675
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
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
         Format          =   20971521
         CurrentDate     =   38056
      End
      Begin VB.Label LblCodigo 
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
         Left            =   150
         TabIndex        =   11
         Top             =   345
         Width           =   570
      End
      Begin VB.Label LblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         TabIndex        =   10
         Top             =   735
         Width           =   495
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
      Left            =   45
      TabIndex        =   8
      Top             =   4305
      Width           =   5820
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   3120
         Picture         =   "FrmMantCompromisoMarcacionUpdate.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cancelar la operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   2205
         Picture         =   "FrmMantCompromisoMarcacionUpdate.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
   End
End
Attribute VB_Name = "FrmMantCompromisoMarcacionUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public lngIdMarcacion As Long
Public datHoraIngresoIdeal As Date
Public datFecha As Date
Public datHoraSalidaIdeal As Date
Public datHoraIngresoReal As Date
Public datHoraSalidaReal As Date

Private datHoraIngresoRealInicio As Date
Private datHoraSalidaRealInicio As Date

Public blnStatus As Boolean
Public strMensajeError As String

Private Const mstrformname As String = "Mantenimiento de Marcación"
Public Sub ShowForm()
On Error GoTo ErrHandler
    Select Case intAction
        Case enumActions.Update
            LoadData
            SettingControls enumActions.Update, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_UPDATE
            datHoraIngresoRealInicio = Me.DtpHoraIngresoReal.Value
            datHoraSalidaRealInicio = Me.DtpHoraSalidaReal.Value
    End Select
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantCompromisoMarcacionUpdate_ShowForm, Err.Source & " FrmMantCompromisoMarcacionUpdate:ShowForm", Err.Description)
End Sub

Private Sub CmdAceptar_Click()
Dim blnResult As Boolean

On Error GoTo ErrHandler
    If Not ValidateFields Then Exit Sub
    If ExisteConsolidado Then
        MsgBox MSG_DATE_CONSOLIDADO_MARCACION, vbCritical, mstrformname
        Exit Sub
    End If
    Select Case intAction
        Case enumActions.Update
            blnResult = Update()
    End Select
    
    If Not blnResult Then
        MsgBox MSG_OPERATION_FAIL & vbCrLf & Me.strMensajeError, vbCritical, mstrformname
        Me.blnStatus = False
    Else
        Unload Me
        If ActiveForm("FrmConsCompromisoMarcacion") Then
            FrmConsCompromisoMarcacion.LoadData
        End If
    End If

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantCompromisoMarcacionUpdate_ShowForm, Err.Source & " FrmMantCompromisoMarcacionUpdate:ShowForm", Err.Description)
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
    Set objPersonal = CreateObject("Personal.clsHorario")
    Set rstPersonal = objPersonal.GetCompromisoMarcacion(, , , , , , , , Me.lngIdMarcacion)
    Set objPersonal = Nothing
    If Not rstPersonal.EOF Then
        Me.txtIdMarcacion.Text = rstPersonal("IdMarcacion")
        Me.TxtIdTrabajador.Text = rstPersonal("IdTrabajador")
        Me.TxtNombres.Text = rstPersonal("Apellido") & " " & rstPersonal("Nombre")
        Me.DTPFecha.Value = rstPersonal("Fecha")
        Me.DtpHoraIngresoIdeal.Value = rstPersonal("HoraIngresoIdeal")
        Me.DtpHoraSalidaIdeal.Value = rstPersonal("HoraSalidaIdeal")
        'If Not IsNull(rstPersonal("HoraIngresoReal")) And Not IsNull(rstPersonal("HoraSalidaReal")) Then
        If rstPersonal("HoraIngresoReal") <> "" And rstPersonal("HoraSalidaReal") <> "" Then
            Me.OptAmbas.Value = True
            Me.DtpHoraIngresoReal.Value = rstPersonal("HoraIngresoReal")
            Me.DtpHoraSalidaReal.Value = rstPersonal("HoraSalidaReal")
        'ElseIf Not IsNull(rstPersonal("HoraIngresoReal")) Then
        ElseIf rstPersonal("HoraIngresoReal") <> "" Then
            Me.OptSoloIngreso.Value = True
            Me.DtpHoraIngresoReal.Value = rstPersonal("HoraIngresoReal")
            'ElseIf not isnull (rstPersonal("HoraSalidaReal"))Then
            ElseIf rstPersonal("HoraSalidaReal") <> "" Then
                Me.OptSoloSalida.Value = True
                Me.DtpHoraSalidaReal.Value = rstPersonal("HoraSalidaReal")
                'ElseIf IsNull(rstPersonal("HoraIngresoReal")) And IsNull(rstPersonal("HoraSalidaReal")) Then
                ElseIf rstPersonal("HoraIngresoReal") = "" And rstPersonal("HoraSalidaReal") = "" Then
                Me.OptNinguno.Value = True
        End If
    End If
    Set rstPersonal = Nothing

Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Set rstPersonal = Nothing
    Call ErrorMessage(FrmMantCompromisoMarcacionUpdate_LoadData, Err.Source & " FrmMantCompromisoMarcacionUpdate:LoadData", Err.Description)
End Sub

Private Function Update() As Boolean
Dim objPersonal As Object
Dim enuOperacion As enumTipoModificacion

On Error GoTo ErrHandler

    Set objPersonal = CreateObject("Personal.clsHorario")
    If Me.OptAmbas.Value Then
        enuOperacion = Ambas
        Update = objPersonal.UpdateMarcacion(Me.lngIdMarcacion, Me.datHoraIngresoIdeal, _
                                    Me.datHoraIngresoReal, Me.datHoraSalidaIdeal, Me.datHoraSalidaReal, enuOperacion)
    ElseIf Me.OptSoloIngreso.Value Then
            enuOperacion = SoloIngreso
            Update = objPersonal.UpdateMarcacion(Me.lngIdMarcacion, Me.datHoraIngresoIdeal, _
                                    Me.datHoraIngresoReal, Me.datHoraSalidaIdeal, , enuOperacion)
        ElseIf Me.OptSoloSalida.Value Then
            enuOperacion = SoloSalida
            Update = objPersonal.UpdateMarcacion(Me.lngIdMarcacion, Me.datHoraIngresoIdeal, _
                                     , Me.datHoraSalidaIdeal, Me.datHoraSalidaReal, enuOperacion)
            ElseIf Me.OptNinguno.Value Then
                enuOperacion = Ninguna
                Update = objPersonal.UpdateMarcacion(Me.lngIdMarcacion, Me.datHoraIngresoIdeal, _
                                     , Me.datHoraSalidaIdeal, , enuOperacion)
            End If
    Set objPersonal = Nothing
    
    If Update Then
        Bitacora "Marcacion", KEY_ACTION_UPDATE, "Codigo " & Str(Me.lngIdMarcacion) & " Ingreso ideal " & Me.datHoraIngresoIdeal & " Salida ideal " & Me.datHoraSalidaIdeal
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Update = False

Call ErrorMessage(FrmMantCompromisoMarcacionUpdate_Update, Err.Source & " FrmMantCompromisoMarcacionUpdate:Update", Err.Description)
End Function

Public Function ValidateFields() As Boolean
    ValidateFields = False

    Me.lngIdMarcacion = Val(Me.txtIdMarcacion.Text)
    Me.datFecha = Me.DTPFecha.Value
    Me.datHoraIngresoIdeal = Format(Me.DtpHoraIngresoIdeal.Value, "hh:mm")
    Me.datHoraSalidaIdeal = Format(Me.DtpHoraSalidaIdeal.Value, "hh:mm")
    Me.datHoraIngresoReal = Format(Me.DtpHoraIngresoReal.Value, "hh:mm")
    Me.datHoraSalidaReal = Format(Me.DtpHoraSalidaReal.Value, "hh:mm")
    ValidateFields = True
End Function

Public Function ExisteConsolidado() As Boolean
Dim objConsolidado As Object
On Error GoTo ErrHandler
    Set objConsolidado = CreateObject("Personal.clsMovimientos")
    ExisteConsolidado = objConsolidado.GetExisteConsolidado(Me.datFecha)
    Set objConsolidado = Nothing

Exit Function
ErrHandler:
    Set objConsolidado = Nothing
    ExisteConsolidado = True
    Call ErrorMessage(FrmMantCompromisoMarcacionUpdate_ExisteConsolidado, Err.Source & " FrmMantCompromisoMarcacionUpdate:ExisteConsolidado", Err.Description)
End Function

