VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMantClientes 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5835
   ClientLeft      =   6420
   ClientTop       =   3975
   ClientWidth     =   5250
   Icon            =   "FrmMantClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   5250
   Begin VB.Frame Frame1 
      Caption         =   "Diferenciar los parámetros de la 1ra marcación del día"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   22
      Top             =   3600
      Width           =   4935
      Begin VB.TextBox txtToleranciaPrimeraMarcacion 
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
         Left            =   2400
         TabIndex        =   25
         Tag             =   "2"
         Top             =   360
         Width           =   1020
      End
      Begin MSComCtl2.DTPicker DtpHoraIngresoReal 
         Height          =   285
         Left            =   2400
         TabIndex        =   23
         Top             =   840
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
         Format          =   63700995
         UpDown          =   -1  'True
         CurrentDate     =   38056
      End
      Begin ComCtl2.UpDown updToleranciaPrimeraMarcacion 
         Height          =   315
         Left            =   3420
         TabIndex        =   28
         Top             =   360
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   327681
         BuddyControl    =   "txtToleranciaPrimeraMarcacion"
         BuddyDispid     =   196633
         OrigLeft        =   3480
         OrigTop         =   240
         OrigRight       =   3720
         OrigBottom      =   615
         Max             =   240
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tolerancia 1ra Entrada :"
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
         TabIndex        =   27
         Top             =   405
         Width           =   2010
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "minutos"
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
         Left            =   3840
         TabIndex        =   26
         Top             =   420
         Width           =   690
      End
      Begin VB.Label LblHoraIngresoReal 
         AutoSize        =   -1  'True
         Caption         =   " Hora Ingreso 1ra Entrada :"
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
         TabIndex        =   24
         Top             =   840
         Width           =   2280
      End
   End
   Begin VB.Frame FraJornada 
      Caption         =   "Tolerancia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   4935
      Begin VB.TextBox txtJorIncompleta 
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
         Left            =   2400
         TabIndex        =   20
         Tag             =   "2"
         Top             =   840
         Width           =   1020
      End
      Begin ComCtl2.UpDown updTardanza 
         Height          =   315
         Left            =   3421
         TabIndex        =   19
         Top             =   285
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   327681
         BuddyControl    =   "txtTardanza"
         BuddyDispid     =   196611
         OrigLeft        =   3480
         OrigTop         =   240
         OrigRight       =   3720
         OrigBottom      =   615
         Max             =   240
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtTardanza 
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
         Left            =   2400
         TabIndex        =   13
         Tag             =   "2"
         Top             =   285
         Width           =   1260
      End
      Begin ComCtl2.UpDown updJorIncompleta 
         Height          =   315
         Left            =   3420
         TabIndex        =   21
         Top             =   840
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   327681
         BuddyControl    =   "txtJorIncompleta"
         BuddyDispid     =   196609
         OrigLeft        =   3600
         OrigTop         =   3240
         OrigRight       =   3840
         OrigBottom      =   3615
         Max             =   240
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "minutos"
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
         Left            =   3840
         TabIndex        =   17
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jornada Incompleta        :"
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
         TabIndex        =   16
         Top             =   885
         Width           =   2115
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "minutos"
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
         Left            =   3840
         TabIndex        =   15
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tolerancia por Tardanza :"
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
         TabIndex        =   14
         Top             =   330
         Width           =   2145
      End
   End
   Begin VB.Frame FraTolerancia 
      Caption         =   "Marcación"
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
      Top             =   1320
      Width           =   4935
      Begin MSComCtl2.UpDown updTolerancia 
         Height          =   315
         Left            =   3421
         TabIndex        =   18
         Top             =   285
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         BuddyControl    =   "txtTolerancia"
         BuddyDispid     =   196617
         OrigLeft        =   3480
         OrigTop         =   240
         OrigRight       =   3720
         OrigBottom      =   615
         Max             =   240
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtTolerancia 
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
         Left            =   2400
         TabIndex        =   9
         Tag             =   "2"
         Top             =   285
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tolerancia Marcación     :"
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
         TabIndex        =   11
         Top             =   330
         Width           =   2070
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "minutos"
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
         Left            =   3840
         TabIndex        =   10
         Top             =   345
         Width           =   690
      End
   End
   Begin VB.OptionButton OptHolgura 
      Caption         =   "Con Holgura"
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
      Left            =   720
      TabIndex        =   6
      Top             =   945
      Width           =   1635
   End
   Begin VB.OptionButton OptMovLibre 
      Caption         =   "Movimiento Libre"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   945
      Width           =   1830
   End
   Begin VB.TextBox txtCliente 
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
      Left            =   1800
      MaxLength       =   60
      TabIndex        =   4
      Top             =   600
      Width           =   3255
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
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   4965
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   2648
         Picture         =   "FrmMantClientes.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cancelar la operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   1733
         Picture         =   "FrmMantClientes.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cliente Asociado"
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
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label LblComentario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Define la tolerancia para considerar en el consolidado de asistencia generado por el sistema."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   360
      TabIndex        =   3
      Top             =   60
      Width           =   4455
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmMantClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public intTiempo As Integer
Public strCliente As String
Public intTolerancia As Integer
Public intTardanza As Integer
Public intJorIncompleta As Integer
Public intTiempoPrimeraMarcacion As Integer
Public datHoraPrimeraMarcacion As Date
Public blnStatus As Boolean
Public strMensajeError As String

Private Const mstrformname As String = "Mantenimiento de Cliente"
Public Sub ShowForm()
Dim objPersonal As Object
Dim rstPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set rstPersonal = objPersonal.GetCliente()
    Set objPersonal = Nothing
    
    If rstPersonal.RecordCount > 0 Then
        intAction = enumActions.Update
    Else
        intAction = enumActions.AddNew
    End If
    Set rstPersonal = Nothing
  
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
    Call ErrorMessage(FrmMantClientes_ShowForm, Err.Source & " FrmMantClientes:ShowForm", Err.Description)
End Sub

Private Sub OptHolgura_Click()
    Me.FraTolerancia.Enabled = True
    Me.Frame1.Enabled = True
    Me.FraJornada.Enabled = True
    Me.txtTolerancia.Enabled = True
    
    Me.Label7.Enabled = True
    Me.Label9.Enabled = True
    Me.intTolerancia = 1
    Me.txtJorIncompleta.Enabled = True
    Me.txtTardanza.Enabled = True
    Me.txtToleranciaPrimeraMarcacion.Enabled = True
    Me.Label1.Enabled = True
    Me.Label2.Enabled = True
    Me.Label3.Enabled = True
    Me.Label4.Enabled = True
    Me.Label6.Enabled = True
    Me.Label8.Enabled = True
    Me.LblHoraIngresoReal.Enabled = True
    Me.DtpHoraIngresoReal.Enabled = True
End Sub

Private Sub OptMovLibre_Click()
    Me.txtTolerancia.Enabled = False
    Me.Label7.Enabled = False
    Me.Label9.Enabled = False
    Me.intTolerancia = 0
    Me.FraTolerancia.Enabled = False
    Me.Frame1.Enabled = False
    Me.FraJornada.Enabled = False
    Me.txtJorIncompleta.Enabled = False
    Me.txtTardanza.Enabled = False
    Me.txtToleranciaPrimeraMarcacion.Enabled = False
    Me.Label1.Enabled = False
    Me.Label2.Enabled = False
    Me.Label3.Enabled = False
    Me.Label4.Enabled = False
    Me.Label6.Enabled = False
    Me.Label8.Enabled = False
    Me.LblHoraIngresoReal.Enabled = False
    Me.DtpHoraIngresoReal.Enabled = False
    
    
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
    End If

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantClientes_ShowForm, Err.Source & " FrmMantClientes:ShowForm", Err.Description)
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
    Set rstPersonal = objPersonal.GetCliente()
    Set objPersonal = Nothing
    If rstPersonal.RecordCount > 0 Then
        'If IsNull(rstPersonal!Cliente) Then
        '    Me.txtCliente.Text = ""
        'Else
            Me.txtCliente.Text = rstPersonal!Cliente
        'End If
        'If IsNull(rstPersonal!Tiempo) Then
        '    Me.txtTolerancia.Text = ""
        'Else
            Me.txtTolerancia.Text = rstPersonal!Tiempo
        'End If
        'If IsNull(rstPersonal!Tiempo) Then
        '    Me.txtTardanza.Text = ""
        'Else
            Me.txtTardanza.Text = rstPersonal!ToleranciaTardanza
        'End If
        'If IsNull(rstPersonal!Tiempo) Then
        '    Me.txtJorIncompleta.Text = ""
        'Else
            Me.txtJorIncompleta.Text = rstPersonal!ToleranciaJornadaIncompleta
        'End If
        'If IsNull(rstPersonal!TiempoPrimeraMarcacion) Then
        '    Me.txtToleranciaPrimeraMarcacion.Text = ""
        'Else
            Me.txtToleranciaPrimeraMarcacion.Text = rstPersonal!TiempoPrimeraMarcacion
        'End If
        'If IsNull(rstPersonal!HoraPrimeraMarcacion) Then
        '    Me.DtpHoraIngresoReal.Value = "00:00"
        'Else
            Me.DtpHoraIngresoReal.Value = rstPersonal!HoraPrimeraMarcacion
        'End If
        If rstPersonal!tolerancia Then
            Me.OptHolgura.Value = True
        Else
            Me.OptMovLibre.Value = True
        End If
    End If
    Set rstPersonal = Nothing

Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Set rstPersonal = Nothing
    Call ErrorMessage(FrmMantClientes_LoadData, Err.Source & " FrmMantClientes:LoadData", Err.Description)
End Sub

Private Function AddNew() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    AddNew = objPersonal.AddNewCliente(Me.strCliente)
    Set objPersonal = Nothing
    If AddNew Then
        Bitacora "Cliente", KEY_ACTION_ADDNEW
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    AddNew = False

Call ErrorMessage(FrmMantClientes_AddNew, Err.Source & " FrmMantClientes:AddNew", Err.Description)
End Function
Private Function Update() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Update = objPersonal.UpdateCliente(Me.intTiempo, Me.strCliente, Me.intTolerancia, Me.intTardanza, Me.intJorIncompleta, Me.intTiempoPrimeraMarcacion, Me.datHoraPrimeraMarcacion)
    Set objPersonal = Nothing
    If Update Then
        Bitacora "Cliente", KEY_ACTION_UPDATE, "Codigo " & Str(Me.intTiempo) & " - Descripcion " & Me.strCliente
        Update = True
    End If
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Update = False

Call ErrorMessage(FrmMantClientes_Update, Err.Source & " FrmMantClientes:Update", Err.Description)
End Function

Public Function ValidateFields() As Boolean
    ValidateFields = False
    If Trim$(txtCliente.Text) = "" Then
        MsgBox MSG_REQUERY_FIELD, vbInformation, mstrformname
        txtCliente.SetFocus
        Exit Function
    End If
    If Me.OptHolgura.Value = True And (txtTolerancia.Text) = "" Then
        MsgBox MSG_REQUERY_FIELD, vbInformation, mstrformname
        txtTolerancia.SetFocus
        Exit Function
    End If
    If Me.OptHolgura.Value = True Then
    If Trim$(txtTardanza.Text) = "" Then
        MsgBox MSG_REQUERY_FIELD, vbInformation, mstrformname
        txtTardanza.SetFocus
        Exit Function
    End If
    If Trim$(txtJorIncompleta.Text) = "" Then
        MsgBox MSG_REQUERY_FIELD, vbInformation, mstrformname
        txtJorIncompleta.SetFocus
        Exit Function
    End If
    End If
    Me.intTiempo = Val(Me.txtTolerancia.Text)
    Me.strCliente = Replace(Me.txtCliente.Text, "'", "''")
    Me.intTardanza = Val(Me.txtTardanza.Text)
    Me.intJorIncompleta = Val(Me.txtJorIncompleta.Text)
    Me.intTiempoPrimeraMarcacion = Val(Me.txtToleranciaPrimeraMarcacion.Text)
    Me.datHoraPrimeraMarcacion = Format(Me.DtpHoraIngresoReal.Value, "hh:mm")
    
    ValidateFields = True
End Function
