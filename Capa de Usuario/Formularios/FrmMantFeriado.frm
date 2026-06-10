VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMantFeriado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Feriado"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "FrmMantFeriado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   3615
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
      Height          =   825
      Left            =   120
      TabIndex        =   6
      Top             =   3375
      Width           =   3360
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   1830
         Picture         =   "FrmMantFeriado.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancelar la operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   900
         Picture         =   "FrmMantFeriado.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Confirmar operación"
         Top             =   195
         Width           =   810
      End
   End
   Begin VB.CheckBox chkRemunerado 
      Caption         =   "Remunerado"
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
      Left            =   330
      TabIndex        =   3
      Top             =   3030
      Width           =   2385
   End
   Begin VB.TextBox txtIdFeriado 
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
      Left            =   1125
      TabIndex        =   0
      Tag             =   "-1"
      Top             =   150
      Width           =   1785
   End
   Begin MSComCtl2.MonthView mvwFecha 
      Height          =   2370
      Left            =   255
      TabIndex        =   2
      Top             =   585
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   21299201
      CurrentDate     =   38404
   End
   Begin VB.Label Label1 
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
      Left            =   315
      TabIndex        =   1
      Top             =   165
      Width           =   660
   End
End
Attribute VB_Name = "FrmMantFeriado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public intIdFeriado As Integer
Public dateFecha As Date
Public blnRemunerado As Integer
Public blnStatus As Boolean
Public strMensajeError As String

Private Const mstrformname As String = "Mantenimiento de Feriado"
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
            Me.txtIdFeriado.Text = 0
            Me.mvwFecha.Value = Date
    End Select
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantFeriado_ShowForm, Err.Source & " FrmMantFeriado:ShowForm", Err.Description)
End Sub

Private Sub CmdAceptar_Click()
Dim blnResult As Boolean

On Error GoTo ErrHandler
    If Not ValidateFields Then Exit Sub
    If ExisteConsolidado Then
        MsgBox MSG_DATE_CONSOLIDADO, vbCritical, mstrformname
        Exit Sub
    End If
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
        If ActiveForm("FrmConsFeriado") Then
            FrmConsFeriado.LoadData
        End If
    End If

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantFeriado_ShowForm, Err.Source & " FrmMantFeriado:ShowForm", Err.Description)
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
    Set rstPersonal = objPersonal.GetFeriado(Me.intIdFeriado)
    Set objPersonal = Nothing
    If Not rstPersonal.EOF Then
        Me.txtIdFeriado.Text = rstPersonal("IdFeriado")
        Me.mvwFecha.Value = rstPersonal("FechaFeriado")
        Me.chkRemunerado.Value = Abs(CInt(rstPersonal("Remunerado")))
    End If
    Set rstPersonal = Nothing

Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Set rstPersonal = Nothing
    Call ErrorMessage(FrmMantFeriado_LoadData, Err.Source & " FrmMantFeriado:LoadData", Err.Description)
End Sub

Private Function AddNew() As Boolean
Dim objPersonal As Object
Dim blnResult As Boolean

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsMovimientos")
    blnResult = objPersonal.AddNewFeriado(Me.dateFecha, Me.blnRemunerado, strMensajeError)
    Set objPersonal = Nothing
    If Not blnResult Then
        MsgBox strMensajeError, vbCritical, mstrformname
        AddNew = False
    Else
        Bitacora "Feriado", KEY_ACTION_ADDNEW
        AddNew = True
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    AddNew = False

Call ErrorMessage(FrmMantFeriado_AddNew, Err.Source & " FrmMantFeriado:AddNew", Err.Description)
End Function
Private Function Update() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsMovimientos")
    Update = objPersonal.UpdateFeriado(Me.intIdFeriado, Me.dateFecha, Me.blnRemunerado)
    Set objPersonal = Nothing
    If Update Then
        Bitacora "Feriado", KEY_ACTION_UPDATE, "Codigo " & Str(Me.intIdFeriado) & " - Fecha " & Me.dateFecha
        Update = True
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Update = False
    Call ErrorMessage(FrmMantFeriado_Update, Err.Source & " FrmMantFeriado:Update", Err.Description)
End Function

Public Function ValidateFields() As Boolean
    ValidateFields = False
    If Trim$(Me.mvwFecha.Value) = "" Then
        MsgBox MSG_REQUERY_FIELD, vbInformation, mstrformname
        Me.mvwFecha.SetFocus
        Exit Function
    End If
    Me.intIdFeriado = Val(Me.txtIdFeriado.Text)
    Me.dateFecha = Me.mvwFecha.Value
    If Me.chkRemunerado.Value Then
        Me.blnRemunerado = 1
    Else
        Me.blnRemunerado = 0
    End If
    
    ValidateFields = True
End Function

Public Function ExisteConsolidado() As Boolean
Dim objConsolidado As Object
On Error GoTo ErrHandler
    Set objConsolidado = CreateObject("Personal.clsMovimientos")
    ExisteConsolidado = objConsolidado.GetExisteConsolidado(Me.dateFecha)
    Set objConsolidado = Nothing
    
Exit Function
ErrHandler:
    Set objConsolidado = Nothing
    ExisteConsolidado = True
    Call ErrorMessage(FrmMantFeriado_ExisteConsolidado, Err.Source & " FrmMantFeriado:ExisteConsolidado", Err.Description)
End Function

