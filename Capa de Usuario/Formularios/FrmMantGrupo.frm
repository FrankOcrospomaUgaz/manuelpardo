VERSION 5.00
Begin VB.Form FrmMantGrupo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   Icon            =   "FrmMantGrupo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5355
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
      Height          =   1140
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   5220
      Begin VB.TextBox TxtIdGrupo 
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
         Left            =   1215
         TabIndex        =   1
         Tag             =   "-1"
         Top             =   300
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
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "2"
         Top             =   690
         Width           =   3915
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
         Left            =   165
         TabIndex        =   7
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
         TabIndex        =   6
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
      Left            =   30
      TabIndex        =   0
      Top             =   1185
      Width           =   5235
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   1913
         Picture         =   "FrmMantGrupo.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   2828
         Picture         =   "FrmMantGrupo.frx":190C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancelar la operación"
         Top             =   210
         Width           =   810
      End
   End
End
Attribute VB_Name = "FrmMantGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public intIdGrupo As Integer
Public strDescripcion As String
Public blnStatus As Boolean
Public strMensajeError As String

Private Const mstrformname As String = "Mantenimiento de Grupo"
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
            Me.TxtIdGrupo.Text = 0
            Me.TxtDescripcion.Text = ""
    End Select
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantGrupo_ShowForm, Err.Source & " FrmMantGrupo:ShowForm", Err.Description)
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
        If ActiveForm("FrmConsGrupo") Then
            FrmConsGrupo.LoadData
        End If
    End If

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantGrupo_ShowForm, Err.Source & " FrmMantGrupo:ShowForm", Err.Description)
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
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set rstPersonal = objPersonal.GetGrupo(Me.intIdGrupo)
    Set objPersonal = Nothing
    If Not rstPersonal.EOF Then
        Me.TxtIdGrupo.Text = rstPersonal("IdGrupo")
        Me.TxtDescripcion.Text = rstPersonal("Descripcion")
    End If
    Set rstPersonal = Nothing

Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Set rstPersonal = Nothing
    Call ErrorMessage(FrmMantGrupo_LoadData, Err.Source & " FrmMantGrupo:LoadData", Err.Description)
End Sub

Private Function AddNew() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    AddNew = objPersonal.AddNewGrupo(Me.strDescripcion)
    Set objPersonal = Nothing
    If AddNew Then
        Bitacora "Grupo", KEY_ACTION_ADDNEW
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    AddNew = False

Call ErrorMessage(FrmMantGrupo_AddNew, Err.Source & " FrmMantGrupo:AddNew", Err.Description)
End Function
Private Function Update() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Update = objPersonal.UpdateGrupo(Me.intIdGrupo, Me.strDescripcion)
    Set objPersonal = Nothing
    If Update Then
        Bitacora "Grupo", KEY_ACTION_UPDATE, "Codigo " & Str(Me.intIdGrupo) & " - Descripcion " & Me.strDescripcion
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Update = False

Call ErrorMessage(FrmMantGrupo_Update, Err.Source & " FrmMantGrupo:Update", Err.Description)
End Function

Public Function ValidateFields() As Boolean
    ValidateFields = False
    If Trim$(TxtDescripcion.Text) = "" Then
        MsgBox MSG_REQUERY_FIELD, vbInformation, mstrformname
        TxtDescripcion.SetFocus
        Exit Function
    End If
    Me.intIdGrupo = Val(Me.TxtIdGrupo.Text)
    Me.strDescripcion = Replace(Me.TxtDescripcion.Text, "'", "''")
    ValidateFields = True
End Function

