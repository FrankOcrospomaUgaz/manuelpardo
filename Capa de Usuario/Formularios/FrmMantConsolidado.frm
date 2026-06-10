VERSION 5.00
Begin VB.Form FrmMantConsolidado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo Consolidado Mensual"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "FrmMantConsolidado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4590
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
      TabIndex        =   7
      Top             =   2175
      Width           =   4215
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   2250
         Picture         =   "FrmMantConsolidado.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancelar la operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   1440
         Picture         =   "FrmMantConsolidado.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
   End
   Begin VB.Frame FraAño 
      Caption         =   "AÑO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   1140
      Width           =   4215
      Begin VB.TextBox txtAño 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ejemplo: 1999, 2004"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2400
         TabIndex        =   6
         Top             =   480
         Width           =   1485
      End
   End
   Begin VB.Frame FraMes 
      Caption         =   "MES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox cboMes 
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
         ItemData        =   "FrmMantConsolidado.frx":1C16
         Left            =   255
         List            =   "FrmMantConsolidado.frx":1C3E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   390
         Width           =   3780
      End
   End
End
Attribute VB_Name = "FrmMantConsolidado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mstrformname As String = "Mantenimiento de Consolidados Mensuales"
Public Sub ShowForm()
    Me.Show
    CenterForm Me
    Me.Caption = mstrformname
End Sub

Private Sub CmdAceptar_Click()
On Error GoTo ErrHandler
    If Len(Trim(Me.txtAño.Text)) <> 0 Then
        If IsNumeric(Me.txtAño.Text) And CInt(Me.txtAño.Text) >= 1000 And CInt(Me.txtAño.Text) <= 9999 Then
            Grabar Me.cboMes.Text, CInt(Me.txtAño.Text)
        Else
            MsgBox "Formato de Año Incorrecto", vbCritical, "Mensaje"
            Me.txtAño.SetFocus
        End If
    Else
        MsgBox "Ingrese un Año"
        Me.txtAño.Text = Trim(Me.txtAño.Text)
    End If

Exit Sub
ErrHandler:
    ErrorMessage FrmMantConsolidado_CmdAceptar_Click, mstrformname & " - CmdAceptar_Click", Err.Description
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub txtAño_GotFocus()
    Me.txtAño.SelStart = 0
    Me.txtAño.SelLength = Len(Me.txtAño.Text)
End Sub

Private Sub txtAño_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Grabar(ByVal Mes As String, ByVal Año As Integer)
Dim objAddNewConsolidado As Object
Dim objBuscarConsolidado As Object
Dim FechaInicio As Date
Dim FechaFin As Date
Dim blnResult As Boolean

On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    Select Case Mes
        Case "ENERO"
            FechaInicio = CDate("01/01/" & Año)
            FechaFin = CDate("31/01/" & Año)
        Case "FEBRERO"
            If Año Mod 4 = 0 Then
                FechaInicio = CDate("01/02/" & Año)
                FechaFin = CDate("29/02/" & Año)
            Else
                FechaInicio = CDate("01/02/" & Año)
                FechaFin = CDate("28/02/" & Año)
            End If
        Case "MARZO"
            FechaInicio = CDate("01/03/" & Año)
            FechaFin = CDate("31/03/" & Año)
        Case "ABRIL"
            FechaInicio = CDate("01/04/" & Año)
            FechaFin = CDate("30/04/" & Año)
        Case "MAYO"
            FechaInicio = CDate("01/05/" & Año)
            FechaFin = CDate("31/05/" & Año)
        Case "JUNIO"
            FechaInicio = CDate("01/06/" & Año)
            FechaFin = CDate("30/06/" & Año)
        Case "JULIO"
            FechaInicio = CDate("01/07/" & Año)
            FechaFin = CDate("31/07/" & Año)
        Case "AGOSTO"
            FechaInicio = CDate("01/08/" & Año)
            FechaFin = CDate("31/08/" & Año)
        Case "SETIEMBRE"
            FechaInicio = CDate("01/09/" & Año)
            FechaFin = CDate("30/09/" & Año)
        Case "OCTUBRE"
            FechaInicio = CDate("01/10/" & Año)
            FechaFin = CDate("31/10/" & Año)
        Case "NOVIEMBRE"
            FechaInicio = CDate("01/11/" & Año)
            FechaFin = CDate("30/11/" & Año)
        Case "DICIEMBRE"
            FechaInicio = CDate("01/12/" & Año)
            FechaFin = CDate("31/12/" & Año)
    End Select
    
    Set objBuscarConsolidado = CreateObject("Personal.clsMovimientos")
    blnResult = objBuscarConsolidado.ExisteMesyAñoConsolidado(Mes, Trim(Str(Año)))
    Set objBuscarConsolidado = Nothing
    If Not blnResult Then
        Dim objform As New FrmMensajeEspera
        CenterForm objform
        objform.Show
        
        Set objAddNewConsolidado = CreateObject("Personal.clsConsolidado")
        blnResult = objAddNewConsolidado.AddNewConsolidado(Mes, Str(Año), FechaInicio, FechaFin)
        Set objAddNewConsolidado = Nothing
        Unload objform
        Set objform = Nothing
        If blnResult Then
            Bitacora "ConsolidadoAsistencia", KEY_ACTION_ADDNEW, Mes & " : " & Str(Año)
            MsgBox "El consolidado se ha grabado correctamente", vbInformation, mstrformname
            Unload Me
            If ActiveForm("FrmConsConsolidadoMensual") Then
                FrmConsConsolidadoMensual.LoadData
            End If
        Else
            MsgBox "Error al crear el consolidado, intente nuevamente", vbCritical, mstrformname
        End If
    Else
        MsgBox "Error... el consolidado existe para el mes y año ingresados", vbExclamation, mstrformname
    End If
    Unload objform
    Set objform = Nothing
    Set objBuscarConsolidado = Nothing
    Set objAddNewConsolidado = Nothing
    Screen.MousePointer = vbDefault

Exit Sub
ErrHandler:
    Set objBuscarConsolidado = Nothing
    Set objAddNewConsolidado = Nothing
    Unload objform
    Set objform = Nothing
    Screen.MousePointer = vbDefault
    ErrorMessage FrmMantConsolidado_Imprimir, mstrformname & " - Imprimir", Err.Description
End Sub
