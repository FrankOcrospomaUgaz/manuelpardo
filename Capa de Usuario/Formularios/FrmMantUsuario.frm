VERSION 5.00
Begin VB.Form FrmMantUsuario 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "FrmMantUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5280
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
      Left            =   120
      TabIndex        =   8
      Top             =   1665
      Width           =   5025
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   2693
         Picture         =   "FrmMantUsuario.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancelar la operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   1800
         Picture         =   "FrmMantUsuario.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
   End
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
      Height          =   1620
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   5025
      Begin VB.TextBox txtClave 
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
         Left            =   1005
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "2"
         Top             =   1170
         Width           =   3975
      End
      Begin VB.TextBox txtNombreUsuario 
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
         Left            =   1005
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "2"
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox TxtIdUsuario 
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
         Left            =   1005
         TabIndex        =   1
         Tag             =   "-1"
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Clave :"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1155
         Width           =   615
      End
      Begin VB.Label LblDescripcion 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
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
         Top             =   750
         Width           =   750
      End
      Begin VB.Label LblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
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
         TabIndex        =   6
         Top             =   315
         Width           =   660
      End
   End
End
Attribute VB_Name = "FrmMantUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public intIdUsuario As Integer
Public strNombreUsuario As String
Public strClave As String
Public blnStatus As Boolean
Public strMensajeError As String
Public strNombreUsuarioAnterior As String
Private Const mstrformname As String = "Mantenimiento de Usuario"
Public Sub ShowForm()
On Error GoTo ErrHandler
    Select Case intAction
        Case enumActions.Update
            LoadData
            SettingControls enumActions.Update, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_UPDATE
            Me.strNombreUsuarioAnterior = Me.txtNombreUsuario.Text
        Case enumActions.AddNew
            Me.strNombreUsuarioAnterior = ""
            LoadData
            SettingControls enumActions.AddNew, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_ADDNEW
            Me.txtIdUsuario.Text = 0
            Me.txtNombreUsuario.Text = ""
            Me.txtClave.Text = ""
    End Select
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantUsuario_ShowForm, Err.Source & " FrmMantUsuario:ShowForm", Err.Description)
End Sub

Private Sub CmdAceptar_Click()
Dim blnResult As Boolean
Dim objPersonal As Object
On Error GoTo ErrHandler
    If Not ValidateFields Then Exit Sub
    Set objPersonal = CreateObject("Personal.clsSeguridad")
    If objPersonal.GetUsuario(, Me.strNombreUsuario, Me.strNombreUsuarioAnterior).RecordCount = 0 Then
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
        If ActiveForm("FrmConsUsuario") Then
            FrmConsUsuario.LoadData
        End If
    End If
    Else
        MsgBox MSG_EXIST_USER, vbInformation, mstrformname
    End If
    Set objPersonal = Nothing
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Call ErrorMessage(FrmMantUsuario_ShowForm, Err.Source & " FrmMantUsuario:ShowForm", Err.Description)
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.strMensajeError = ""
End Sub

Private Sub LoadData()
Dim objSeguridad As Object
Dim rstSeguridad As Object

On Error GoTo ErrHandler
    Set objSeguridad = CreateObject("Personal.clsSeguridad")
    Set rstSeguridad = objSeguridad.GetUsuario(Me.intIdUsuario)
    Set objSeguridad = Nothing
    If Not rstSeguridad.EOF Then
        Me.txtIdUsuario.Text = rstSeguridad("IdUsuario")
        Me.txtNombreUsuario.Text = rstSeguridad("NombreUsuario")
        Me.txtClave.Text = rstSeguridad("Clave")
    End If
    Set rstSeguridad = Nothing

Exit Sub
ErrHandler:
    Set objSeguridad = Nothing
    Set rstSeguridad = Nothing
    Call ErrorMessage(FrmMantUsuario_LoadData, Err.Source & " FrmMantUsuario:LoadData", Err.Description)
End Sub

Private Function AddNew() As Boolean
Dim objSeguridad As Object
On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    Set objSeguridad = CreateObject("Personal.clsSeguridad")
    AddNew = objSeguridad.AddNewUsuario(Me.strNombreUsuario, Me.strClave)
    Set objSeguridad = Nothing
    If AddNew Then
        Bitacora "Usuario", KEY_ACTION_ADDNEW
    End If
    Screen.MousePointer = vbDefault
Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    Set objSeguridad = Nothing
    AddNew = False
Call ErrorMessage(FrmMantUsuario_AddNew, Err.Source & " FrmMantUsuario : AddNew", Err.Description)
End Function
Private Function Update() As Boolean
Dim objSeguridad As Object

On Error GoTo ErrHandler
    Set objSeguridad = CreateObject("Personal.clsSeguridad")
    
    Update = objSeguridad.UpdateUsuario(Me.intIdUsuario, Me.strNombreUsuario, Me.strClave)
    Set objSeguridad = Nothing
    If Update Then
        Bitacora "Usuario", KEY_ACTION_UPDATE, "Codigo " & Str(Me.intIdUsuario) & " - NombreUsuario " & Me.strNombreUsuario
    End If
    
Exit Function
ErrHandler:
    Set objSeguridad = Nothing
    Update = False

Call ErrorMessage(FrmMantUsuario_Update, Err.Source & " FrmMantUsuario:Update", Err.Description)
End Function

Public Function ValidateFields() As Boolean
    ValidateFields = False
    If Trim$(txtNombreUsuario.Text) = "" Then
        MsgBox MSG_REQUERY_FIELD, vbInformation, mstrformname
        txtNombreUsuario.SetFocus
        Exit Function
    End If
    Me.intIdUsuario = Val(Me.txtIdUsuario.Text)
    Me.strNombreUsuario = Replace(Me.txtNombreUsuario.Text, "'", "''")
    Me.strClave = Replace(Me.txtClave.Text, "'", "''")
    ValidateFields = True
End Function


