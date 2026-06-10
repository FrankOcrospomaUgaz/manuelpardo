VERSION 5.00
Begin VB.Form FrmMantTipoHora 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "FrmMantTipoHora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5775
   Begin VB.Frame FraDatos 
      Caption         =   "Datos a registrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   75
      TabIndex        =   7
      Top             =   45
      Width           =   5655
      Begin VB.OptionButton OptNo 
         Caption         =   "No"
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
         Left            =   3165
         TabIndex        =   4
         Top             =   1155
         Width           =   1695
      End
      Begin VB.OptionButton OptSi 
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1395
         TabIndex        =   3
         Top             =   1155
         Width           =   1485
      End
      Begin VB.TextBox TxtIdTipoHora 
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
         Left            =   1350
         TabIndex        =   1
         Tag             =   "-1"
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox TxtDescripcion 
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
         Left            =   1350
         MaxLength       =   20
         TabIndex        =   2
         Tag             =   "2"
         Top             =   705
         Width           =   4170
      End
      Begin VB.Label LblRemunerada 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   1155
         Width           =   1095
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
         TabIndex        =   9
         Top             =   330
         Width           =   570
      End
      Begin VB.Label LblDescripcion 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
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
         TabIndex        =   8
         Top             =   735
         Width           =   975
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
      Left            =   75
      TabIndex        =   0
      Top             =   1665
      Width           =   5655
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   1995
         Picture         =   "FrmMantTipoHora.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   2925
         Picture         =   "FrmMantTipoHora.frx":190C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cancelar la operación"
         Top             =   225
         Width           =   810
      End
   End
End
Attribute VB_Name = "FrmMantTipoHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public intIdTipoHora As Integer
Public strDescripcion As String
Public strRemunerada As String
Public blnStatus As Boolean
Public strMensajeError As String

Private Const mstrformname As String = "Mantenimiento de Tipo de Hora"

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
            Me.TxtIdTipoHora.Text = 0
            Me.TxtDescripcion.Text = ""
            Me.OptSi.Value = True
    End Select
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantTipoHora_ShowForm, Err.Source & " FrmMantTipoHora:ShowForm", Err.Description)
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
        If ActiveForm("FrmConsTipoHora") Then
            FrmConsTipoHora.LoadData
        End If
    End If

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantTipoHora_ShowForm, Err.Source & " FrmMantTipoHora:ShowForm", Err.Description)
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

'On Error GoTo ErrHandler
'    Set objPersonal = CreateObject("Personal.clsPersonal")
'    Set rstPersonal = objPersonal.GetTipoHora(Me.intIdTipoHora)
'    Set objPersonal = Nothing
'    If Not rstPersonal.EOF Then
'        Me.TxtIdTipoHora.Text = rstPersonal("IdTipoHora")
'        Me.TxtDescripcion.Text = rstPersonal("Descripcion")
'        If rstPersonal("Remunerada") = gstrHoraRemunerada Then
'            Me.OptSi.Value = True
'        Else
'            Me.OptNo.Value = True
'        End If
'    End If
'    Set rstPersonal = Nothing

Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Set rstPersonal = Nothing
    Call ErrorMessage(FrmMantTipoHora_LoadData, Err.Source & " FrmMantTipoHora:LoadData", Err.Description)
End Sub

Private Function AddNew() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    AddNew = objPersonal.AddNewTipoHora(Me.strDescripcion, Me.strRemunerada)
    Set objPersonal = Nothing
    If AddNew Then
        Bitacora "TipoHora", KEY_ACTION_ADDNEW
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    AddNew = False

Call ErrorMessage(FrmMantTipoHora_AddNew, Err.Source & " FrmMantTipoHora:AddNew", Err.Description)
End Function
Private Function Update() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Update = objPersonal.UpdateTipoHora(Me.intIdTipoHora, Me.strDescripcion, Me.strRemunerada)
    Set objPersonal = Nothing
    If Update Then
        Bitacora "TipoHora", KEY_ACTION_UPDATE, "Codigo " & Str(Me.intIdTipoHora) & " - Descripcion " & Me.strDescripcion
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Update = False

Call ErrorMessage(FrmMantTipoHora_Update, Err.Source & " FrmMantTipoHora:Update", Err.Description)
End Function

Public Function ValidateFields() As Boolean
'    ValidateFields = False
'    If Trim$(TxtDescripcion.Text) = "" Then
'        MsgBox MSG_REQUERY_FIELD, vbInformation, mstrformname
'        TxtDescripcion.SetFocus
'        Exit Function
'    End If
'    Me.intIdTipoHora = Val(Me.TxtIdTipoHora.Text)
'    Me.strDescripcion = Replace(Me.TxtDescripcion.Text, "'", "''")
'    If Me.OptSi.Value = True Then
'        Me.strRemunerada = gstrHoraRemunerada
'    Else
'        Me.strRemunerada = gstrHoraNoRemunerada
'    End If
'    ValidateFields = True
End Function



