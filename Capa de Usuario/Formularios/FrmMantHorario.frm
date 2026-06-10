VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMantHorario 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "FrmMantHorario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5520
   Begin VB.Frame FraDatos 
      Caption         =   "Datos a registrar"
      Height          =   1860
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5265
      Begin MSComCtl2.DTPicker DtpHoraSalida 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   21299202
         CurrentDate     =   38056
      End
      Begin MSComCtl2.DTPicker DtpHoraIngreso 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   21299202
         CurrentDate     =   38056
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   21299201
         CurrentDate     =   38056
      End
      Begin VB.TextBox TxtIdHorario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1455
         TabIndex        =   4
         Tag             =   "-1"
         Top             =   300
         Width           =   855
      End
      Begin VB.Label LblHoraSalida 
         AutoSize        =   -1  'True
         Caption         =   "Hora de Salida"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label LblHoraIngreso 
         AutoSize        =   -1  'True
         Caption         =   "Hora de Ingreso"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label LblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   330
         Width           =   495
      End
      Begin VB.Label LblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   735
         Width           =   450
      End
   End
   Begin VB.Frame FraAcciones 
      Caption         =   "Acciones"
      Height          =   810
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   5235
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   1913
         Picture         =   "FrmMantHorario.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   2828
         Picture         =   "FrmMantHorario.frx":190C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cancelar la operación"
         Top             =   210
         Width           =   810
      End
   End
End
Attribute VB_Name = "FrmMantHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public intIdHorario As Long
Public datFecha As Date
Public datHoraIngreso As Date
Public datHoraSalida As Date
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
        Case enumActions.AddNew
            LoadData
            SettingControls enumActions.AddNew, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_ADDNEW
            Me.TxtIdHorario.Text = 0
            Me.DTPFecha.Value = Date
            Me.DtpHoraIngreso.Value = Time
            Me.DtpHoraSalida.Value = Time
    End Select
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantHorario_ShowForm, Err.Source & " FrmMantHorario:ShowForm", Err.Description)
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
        If ActiveForm("FrmConsHorario") Then
            FrmConsHorario.LoadData
        End If
    End If

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantHorario_ShowForm, Err.Source & " FrmMantHorario:ShowForm", Err.Description)
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
    Set rstPersonal = objPersonal.GetHorario(Me.intIdHorario)
    Set objPersonal = Nothing
    If Not rstPersonal.EOF Then
        Me.TxtIdHorario.Text = rstPersonal("IdHorario")
        Me.DTPFecha.Value = rstPersonal("Fecha")
        Me.DtpHoraIngreso.Value = rstPersonal("HoraIngreso")
        Me.DtpHoraSalida.Value = rstPersonal("HoraSalida")
    End If
    Set rstPersonal = Nothing

Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Set rstPersonal = Nothing
    Call ErrorMessage(FrmMantHorario_LoadData, Err.Source & " FrmMantHorario:LoadData", Err.Description)
End Sub

Private Function AddNew() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsHorario")
    AddNew = objPersonal.AddNewHorario(Me.datFecha, Me.datHoraIngreso, Me.datHoraSalida)
    Set objPersonal = Nothing
    If AddNew Then
        Bitacora "Horario", KEY_ACTION_ADDNEW
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    AddNew = False

Call ErrorMessage(FrmMantHorario_AddNew, Err.Source & " FrmMantHorario:AddNew", Err.Description)
End Function
Private Function Update() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsHorario")
    Update = objPersonal.UpdateHorario(Me.intIdHorario, Me.datFecha, Me.datHoraIngreso, Me.datHoraSalida)
    Set objPersonal = Nothing
    If Update Then
        Bitacora "Horario", KEY_ACTION_UPDATE, "Codigo " & Str(Me.intIdHorario) & " - Fecha " & Me.datFecha
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Update = False

Call ErrorMessage(FrmMantHorario_Update, Err.Source & " FrmMantHorario:Update", Err.Description)
End Function

Public Function ValidateFields() As Boolean
    ValidateFields = False
    If Me.DtpHoraIngreso.Value - Me.DtpHoraSalida.Value >= 0 Then
        MsgBox MSG_INVALIDATE_KEY, vbInformation, mstrformname
        DtpHoraIngreso.SetFocus
        Exit Function
    End If
    Me.intIdHorario = Val(Me.TxtIdHorario.Text)
    Me.datFecha = Me.DTPFecha.Value
    Me.datHoraIngreso = Me.DtpHoraIngreso.Value
    Me.datHoraSalida = Me.DtpHoraSalida.Value
    ValidateFields = True
End Function



