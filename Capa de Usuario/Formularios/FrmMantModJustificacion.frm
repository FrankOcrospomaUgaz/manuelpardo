VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMantModJustificacion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "FrmMantModJustificacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6975
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
      Left            =   0
      TabIndex        =   11
      Top             =   2520
      Width           =   6795
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   2565
         Picture         =   "FrmMantModJustificacion.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   3480
         Picture         =   "FrmMantModJustificacion.frx":190C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cancelar la operación"
         Top             =   210
         Width           =   810
      End
   End
   Begin VB.Frame FraDatos 
      Caption         =   "Datos de Justificación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6795
      Begin VB.TextBox TxtIdJustificacion 
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
         Left            =   1230
         TabIndex        =   4
         Tag             =   "-1"
         Top             =   540
         Width           =   1560
      End
      Begin VB.TextBox TxtObservacion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1230
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "FrmMantModJustificacion.frx":1C16
         Top             =   1680
         Width           =   5370
      End
      Begin VB.OptionButton OptNoRemunerada 
         Caption         =   "No Remunerada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3562
         TabIndex        =   2
         Top             =   210
         Width           =   1740
      End
      Begin VB.OptionButton OptRemunerada 
         Caption         =   "Remunerada"
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
         Left            =   1552
         TabIndex        =   1
         Top             =   210
         Width           =   1395
      End
      Begin MSDataListLib.DataCombo DtcMotivoJustificacion 
         Height          =   315
         Left            =   1230
         TabIndex        =   5
         Top             =   915
         Width           =   3960
         _ExtentX        =   6985
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
      Begin MSComCtl2.DTPicker DTPFecha 
         Height          =   285
         Left            =   1230
         TabIndex        =   6
         Top             =   1305
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   21299201
         CurrentDate     =   38056
      End
      Begin VB.Label LblMotivo 
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
         Left            =   135
         TabIndex        =   10
         Top             =   945
         Width           =   585
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
         Left            =   105
         TabIndex        =   9
         Top             =   555
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
         Left            =   135
         TabIndex        =   8
         Top             =   1350
         Width           =   495
      End
      Begin VB.Label LblObservacion 
         AutoSize        =   -1  'True
         Caption         =   "Observación"
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
         TabIndex        =   7
         Top             =   1680
         Width           =   1050
      End
   End
End
Attribute VB_Name = "FrmMantModJustificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intIdJustificacion As Integer
Public intIdMotivoJustificacion As String
Public datFecha As String
Public strObservacion As String
Public strRemunerada As String

Private mrstMotivoJustificacion As Object

Public blnStatus As Boolean
Public strMensajeError As String

Private Const mstrformname As String = "Mantenimiento de Justificaciones"

Public Property Get rstMotivoJustificacion() As Variant
    Set rstMotivoJustificacion = mrstMotivoJustificacion
End Property
Public Property Set rstMotivoJustificacion(ByVal vNewValue As Variant)
    Set mrstMotivoJustificacion = vNewValue
    If Not rstMotivoJustificacion.EOF Then
        Set Me.DtcMotivoJustificacion.RowSource = rstMotivoJustificacion
        Me.DtcMotivoJustificacion.ListField = "Descripcion"
        Me.DtcMotivoJustificacion.BoundColumn = "IdMotivoJustificacion"
        Me.DtcMotivoJustificacion.Enabled = True
    Else
        Me.DtcMotivoJustificacion.Enabled = False
        Me.DtcMotivoJustificacion.Text = ""
    End If
End Property

Public Sub ShowForm()
On Error GoTo ErrHandler
    LoadData
    SettingControls enumActions.Update, Me
    Me.Caption = mstrformname & " - " & KEY_ACTION_UPDATE
    
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    
    Call ErrorMessage(FrmMantModJustificacion_ShowForm, Err.Source & " FrmMantModJustificacion:ShowForm", Err.Description)
End Sub

Private Sub CmdAceptar_Click()
Dim blnResult As Boolean
On Error GoTo ErrHandler
    If Not ValidateFields Then Exit Sub
        blnResult = Update()

    If Not blnResult Then
        MsgBox MSG_OPERATION_FAIL & vbCrLf & Me.strMensajeError, vbCritical, mstrformname
        Me.blnStatus = False
    Else
        Unload Me
        If ActiveForm("FrmConsJustificacion") Then
            FrmConsJustificacion.LoadData
        End If
        FrmConsJustificacion.Show
    End If

Exit Sub
ErrHandler:
Call ErrorMessage(FrmMantModJustificacion_ShowForm, Err.Source & " FrmMantModJustificacion:ShowForm", Err.Description)
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
    If ActiveForm("FrmConsJustificacion") Then
            FrmConsJustificacion.LoadData
    End If
    FrmConsJustificacion.Show
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.strMensajeError = ""
    Me.Caption = mstrformname
End Sub

Private Sub LoadData()
Dim objPersonal As Object
Dim rstPersonal As Object
    
On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsMovimientos")
    Set rstPersonal = objPersonal.GetJustificacion(Me.intIdJustificacion, , , , gdatFecha_Nula_RangoInicio, gdatFecha_Nula_RangoFin)
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set rstMotivoJustificacion = objPersonal.GetMotivoJustificacion()
    Set objPersonal = Nothing
    
    If Not rstPersonal.EOF Then
        Me.TxtIdJustificacion.Text = rstPersonal("IdJustificacion")
        Me.DtcMotivoJustificacion.BoundText = rstPersonal("IdMotivoJustificacion")
        Me.DTPFecha.Value = rstPersonal("Fecha")
        Me.TxtObservacion.Text = rstPersonal("Observacion")
        If rstPersonal("Observacion") = "N" Then
            Me.OptRemunerada.Value = 1
            Me.OptNoRemunerada.Value = 0
        Else
            Me.OptRemunerada.Value = 0
            Me.OptNoRemunerada.Value = 1
        End If
    End If
    Set rstPersonal = Nothing
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Set rstPersonal = Nothing
    Call ErrorMessage(FrmMantModJustificacion_LoadData, Err.Source & " FrmMantModJustificacion:LoadData", Err.Description)
End Sub

Private Function Update() As Boolean
Dim objPersonal As Object
Dim objGeneral As Object
Dim strMensajeError As String
Dim blnResult As Boolean

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsMovimientos")
    blnResult = objPersonal.UpdateJustificacion(Me.intIdJustificacion, Me.intIdMotivoJustificacion, Me.strRemunerada, Me.strObservacion, Me.datFecha)
    Set objPersonal = Nothing
    If Not blnResult Then
        MsgBox strMensajeError, vbCritical, mstrformname
        Update = False
    Else
        Update = True
        Set objGeneral = CreateObject("Personal.clsGeneral")
        objGeneral.AddNewBitacora gstrNombreUsuario, "Justificacion ", KEY_ACTION_UPDATE, ""
        Set objGeneral = Nothing
    End If

Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Set objGeneral = Nothing
    Update = False

Call ErrorMessage(FrmMantModJustificacion_UpDate, Err.Source & " FrmMantModJustificacion:Update", Err.Description)
End Function

Public Function ValidateFields() As Boolean
    ValidateFields = False

    If Me.OptNoRemunerada.Value = 0 And Me.OptRemunerada.Value = 0 Then
        MsgBox "Debe seleccionar una opcion", vbInformation, mstrformname
        Exit Function
    End If
    If Trim$(Me.DtcMotivoJustificacion.Text) = "" Then
        MsgBox "Debe seleccionar una categoría", vbInformation, mstrformname
        Me.DtcMotivoJustificacion.SetFocus
        Exit Function
    End If

    Me.intIdJustificacion = Val(Me.TxtIdJustificacion.Text)
    Me.intIdMotivoJustificacion = Me.DtcMotivoJustificacion.BoundText
    If Me.OptRemunerada.Value Then
        Me.strRemunerada = "S"
    Else
        Me.strRemunerada = "N"
    End If
    Me.strObservacion = Replace(Me.TxtObservacion.Text, "'", "''")
    Me.datFecha = Me.DTPFecha.Value
    
    ValidateFields = True
End Function

Private Sub Form_Terminate()
    Set mrstMotivoJustificacion = Nothing
End Sub
