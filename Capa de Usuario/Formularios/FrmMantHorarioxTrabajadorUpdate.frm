VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMantHorarioxTrabajadorUpdate 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "FrmMantHorarioxTrabajadorUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5955
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
      Left            =   60
      TabIndex        =   9
      Top             =   3165
      Width           =   5820
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   2205
         Picture         =   "FrmMantHorarioxTrabajadorUpdate.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   3120
         Picture         =   "FrmMantHorarioxTrabajadorUpdate.frx":190C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cancelar la operación"
         Top             =   210
         Width           =   810
      End
   End
   Begin VB.Frame FraHorario 
      Caption         =   "Datos de Horario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   30
      TabIndex        =   12
      Top             =   1080
      Width           =   5820
      Begin VB.TextBox TxtIdHorario 
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
         TabIndex        =   3
         Tag             =   "-1"
         Top             =   315
         Width           =   1620
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
         TabIndex        =   13
         Top             =   1110
         Width           =   5520
         Begin MSComCtl2.DTPicker DtpHoraSalidaIdeal 
            Height          =   285
            Left            =   4035
            TabIndex        =   6
            Tag             =   "2"
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
            Format          =   63438851
            UpDown          =   -1  'True
            CurrentDate     =   38056
         End
         Begin MSComCtl2.DTPicker DtpHoraIngresoIdeal 
            Height          =   285
            Left            =   960
            TabIndex        =   5
            Tag             =   "2"
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
            Format          =   63438851
            UpDown          =   -1  'True
            CurrentDate     =   38056
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
            TabIndex        =   15
            Top             =   375
            Width           =   510
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
            TabIndex        =   14
            Top             =   375
            Width           =   705
         End
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   300
         Left            =   1065
         TabIndex        =   4
         Tag             =   "2"
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
         Format          =   63438849
         CurrentDate     =   38056
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
         TabIndex        =   17
         Top             =   735
         Width           =   495
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
         TabIndex        =   16
         Top             =   345
         Width           =   570
      End
   End
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
      TabIndex        =   0
      Top             =   15
      Width           =   5835
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
         TabIndex        =   2
         Tag             =   "-1"
         Top             =   675
         Width           =   4635
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
         Left            =   1080
         TabIndex        =   1
         Tag             =   "-1"
         Text            =   "0"
         Top             =   315
         Width           =   1560
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
         TabIndex        =   11
         Top             =   660
         Width           =   750
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
         TabIndex        =   10
         Top             =   300
         Width           =   570
      End
   End
End
Attribute VB_Name = "FrmMantHorarioxTrabajadorUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public intIdHorario As Long
Public datHoraIngresoIdeal As Date
Public datHoraSalidaIdeal As Date
Public datFecha As Date

Public blnStatus As Boolean
Public strMensajeError As String

Private Const mstrformname As String = "Mantenimiento de Horario"
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
    Call ErrorMessage(FrmMantHorarioxTrabajadorUpdate_ShowForm, Err.Source & " FrmMantHorarioxTrabajadorUpdate:ShowForm", Err.Description)
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
        Me.blnStatus = False
    Else
        Unload Me
        If ActiveForm("FrmConsHorarioxTrabajador") Then
            FrmConsHorarioxTrabajador.LoadData
        End If
    End If

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantHorarioxTrabajadorUpdate_ShowForm, Err.Source & " FrmMantHorarioxTrabajadorUpdate:ShowForm", Err.Description)
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
    Set rstPersonal = objPersonal.GetHorarioxTrabajador(, , , , , , , , , , Me.intIdHorario)
    Set objPersonal = Nothing
    If Not rstPersonal.EOF Then
        Me.TxtIdHorario.Text = rstPersonal("IdHorario")
        Me.TxtIdTrabajador.Text = rstPersonal("IdTrabajador")
        Me.TxtNombres.Text = rstPersonal("Apellido") & " " & rstPersonal("Nombre")
        Me.DTPFecha.Value = rstPersonal("Fecha")
        Me.DtpHoraIngresoIdeal.Value = rstPersonal("HoraIngreso")
        Me.DtpHoraSalidaIdeal.Value = rstPersonal("HoraSalida")
    End If
    Set rstPersonal = Nothing

Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Set rstPersonal = Nothing
    Call ErrorMessage(FrmMantHorarioxTrabajadorUpdate_LoadData, Err.Source & " FrmMantHorarioxTrabajadorUpdate:LoadData", Err.Description)
End Sub

Private Function Update() As Boolean
Dim objPersonal As Object
Dim strMensajeError As String
Dim blnResult As Boolean

On Error GoTo ErrHandler

    Set objPersonal = CreateObject("Personal.clsHorario")
    blnResult = objPersonal.UpdateHorarioxTrabajador(Me.intIdHorario, Me.datFecha, _
                                    Me.datHoraIngresoIdeal, Me.datHoraSalidaIdeal, strMensajeError)
    Set objPersonal = Nothing
    If Not blnResult Then
        MsgBox strMensajeError, vbCritical, mstrformname
        Update = False
    Else
        Update = True
        Bitacora "HorarioxTrabajador", KEY_ACTION_UPDATE, "Codigo " & Str(Me.intIdHorario) & " Ingreso ideal " & Me.datHoraIngresoIdeal & " Salida ideal " & Me.datHoraSalidaIdeal
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Update = False

Call ErrorMessage(FrmMantHorarioxTrabajadorUpdate_Update, Err.Source & " FrmMantHorarioxTrabajadorUpdate:Update", Err.Description)
End Function

Public Function ValidateFields() As Boolean
    ValidateFields = False

    Me.intIdHorario = Val(Me.TxtIdHorario.Text)
    Me.datFecha = Me.DTPFecha.Value
    Me.datHoraIngresoIdeal = FormatDateTime(Me.DtpHoraIngresoIdeal.Value, vbShortTime) '= Format(Me.DtpHoraIngresoIdeal.Value, "hh:mm")
    Me.datHoraSalidaIdeal = FormatDateTime(Me.DtpHoraSalidaIdeal.Value, vbShortTime)   ' = Format(Me.DtpHoraSalidaIdeal.Value, "hh:mm")
    ValidateFields = True
End Function
