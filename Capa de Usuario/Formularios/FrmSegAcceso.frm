VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Begin VB.Form FrmSegAcceso 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso al Sistema "
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   Icon            =   "FrmSegAcceso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   3480
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   795
      Left            =   5895
      Picture         =   "FrmSegAcceso.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1935
      Width           =   960
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   795
      Left            =   5895
      Picture         =   "FrmSegAcceso.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1035
      Width           =   960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1770
      Left            =   1530
      TabIndex        =   0
      Top             =   990
      Width           =   4200
      Begin VB.TextBox txtClave 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   900
         Width           =   2940
      End
      Begin VB.TextBox txtUsuario 
         Height          =   330
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   2940
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contraseña :"
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   990
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuario :"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   405
         Width           =   630
      End
   End
   Begin VB.Image ImgLogo 
      Height          =   900
      Left            =   5895
      Picture         =   "FrmSegAcceso.frx":149E
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de Control de Personal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1530
      TabIndex        =   7
      Top             =   105
      Width           =   4425
   End
   Begin VB.Image Image1 
      Height          =   3090
      Left            =   0
      Picture         =   "FrmSegAcceso.frx":1E4A
      Stretch         =   -1  'True
      Top             =   -90
      Width           =   1455
   End
End
Attribute VB_Name = "FrmSegAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strUsuario As String
Public strClave As String

Private Const mstrformname As String = "Ingreso al Sistema"

Private Sub CmdAceptar_Click()
Dim objPersonal As Object
Dim rstUsuario As Object
Dim rstMenuUsuario As Object
Dim strAccesoMenu As String
Dim i As Integer
On Error GoTo ErrHandler
    strAccesoMenu = ""
    If ValidateFields Then
        Screen.MousePointer = vbHourglass
        Set objPersonal = CreateObject("Personal.clsSeguridad")
        Set rstUsuario = objPersonal.GetValidarUsuario(Me.strUsuario, Me.strClave)
        Set objPersonal = Nothing
        
        If rstUsuario.RecordCount > 0 Then
            gstrNombreUsuario = UCase$(Me.strUsuario)
            'aqui debe configurar el menu de sistemas para el usuario
'            Set objPersonal = CreateObject("Personal.clsSeguridad")
'            Set rstMenuUsuario = objPersonal.GetOpcionesModulo(strSistema)
'            Do While Not rstMenuUsuario.EOF
'                rstMenuUsuario!PropiedadName.Visible = True
'            Loop
            
            'configura el menu del modulo seleccionado
            Set objPersonal = CreateObject("Personal.clsSeguridad")
            Set rstMenuUsuario = objPersonal.GetAccesoUnicoUsuario(CInt(rstUsuario!IdUsuario))
            Set objPersonal = Nothing
            
            Do While Not rstMenuUsuario.EOF
                If rstMenuUsuario!Habilitado Then
                    strAccesoMenu = strAccesoMenu & "1"
                Else
                    strAccesoMenu = strAccesoMenu & "0"
                End If
                rstMenuUsuario.MoveNext
            Loop
            
            '*******    Para Inhabilitar los Menús Del Usuario Entrante
            MDIPrincipal.mnuArea.Enabled = IIf(Mid(strAccesoMenu, MENU_MANTENIMIENTO_AREA, 1) = "1", True, False)
            MDIPrincipal.mnuCargo.Enabled = IIf(Mid(strAccesoMenu, MENU_MANTENIMIENTO_CARGO, 1) = "1", True, False)
            MDIPrincipal.mnuCategoria.Enabled = IIf(Mid(strAccesoMenu, MENU_MANTENIMIENTO_CATEGORIA, 1) = "1", True, False)
            MDIPrincipal.mnuTipoHora.Enabled = IIf(Mid(strAccesoMenu, MENU_MANTENIMIENTO_TIPO_HORA, 1) = "1", True, False)
            MDIPrincipal.mnuMotivoJustificacion.Enabled = IIf(Mid(strAccesoMenu, MENU_MANTENIMIENTO_MOTIVO_JUSTIFICACION, 1) = "1", True, False)
            MDIPrincipal.mnuFeriado.Enabled = IIf(Mid(strAccesoMenu, MENU_MANTENIMIENTO_FERIADO, 1) = "1", True, False)
            MDIPrincipal.mnuGrupoTrabajador.Enabled = IIf(Mid(strAccesoMenu, MENU_MANTENIMIENTO_GRUPO, 1) = "1", True, False)
            MDIPrincipal.mnuTrabajador.Enabled = IIf(Mid(strAccesoMenu, MENU_MANTENIMIENTO_TRABAJADOR, 1) = "1", True, False)
            MDIPrincipal.mnuCliente.Enabled = IIf(Mid(strAccesoMenu, MENU_MANTENIMIENTO_CLIENTE, 1) = "1", True, False)
            MDIPrincipal.mnuUsuarios.Enabled = IIf(Mid(strAccesoMenu, MENU_MANTENIMIENTO_USUARIO, 1) = "1", True, False)
            MDIPrincipal.mnuDefinirAccesoUsuario.Enabled = IIf(Mid(strAccesoMenu, MENU_MANTENIMIENTO_ACCESO_USUARIO, 1) = "1", True, False)
            MDIPrincipal.mnuDefinirHorarios.Enabled = IIf(Mid(strAccesoMenu, MENU_MOVIMIENTO_DEFINIR_HORARIO, 1) = "1", True, False)
            MDIPrincipal.mnuJustificaciones.Enabled = IIf(Mid(strAccesoMenu, MENU_MOVIMIENTO_JUSTIFICACION, 1) = "1", True, False)
            MDIPrincipal.mnuPermisos.Enabled = IIf(Mid(strAccesoMenu, MENU_MOVIMIENTO_PERMISO, 1) = "1", True, False)
            MDIPrincipal.mnuSuspensiones.Enabled = IIf(Mid(strAccesoMenu, MENU_MOVIMIENTO_SUSPENSION, 1) = "1", True, False)
            MDIPrincipal.mnuGenerarCompromisosMarcacion.Enabled = IIf(Mid(strAccesoMenu, MENU_MOVIMIENTO_GENERAR_COMP_MARC, 1) = "1", True, False)
            MDIPrincipal.mnuMarcacion.Enabled = IIf(Mid(strAccesoMenu, MENU_MOVIMIENTO_MARCACION, 1) = "1", True, False)
            MDIPrincipal.mnuConsAsistMens.Enabled = IIf(Mid(strAccesoMenu, MENU_MOVIMIENTO_CONSOLIDADO_MARCACION, 1) = "1", True, False)
            MDIPrincipal.mnuListaTrab.Enabled = IIf(Mid(strAccesoMenu, MENU_REPORTE_TRABAJADORES, 1) = "1", True, False)
            MDIPrincipal.mnuListadoHor.Enabled = IIf(Mid(strAccesoMenu, MENU_REPORTE_HORARIOS, 1) = "1", True, False)
            MDIPrincipal.mnuRegMarc.Enabled = IIf(Mid(strAccesoMenu, MENU_REPORTE_MARCACIONES, 1) = "1", True, False)
            MDIPrincipal.mnuComparacion.Enabled = IIf(Mid(strAccesoMenu, MENU_REPORTE_COMPARACION, 1) = "1", True, False)
            MDIPrincipal.mnuListaSusp.Enabled = IIf(Mid(strAccesoMenu, MENU_REPORTE_SUSPENSIONES, 1) = "1", True, False)
            MDIPrincipal.mnuListJust.Enabled = IIf(Mid(strAccesoMenu, MENU_REPORTE_JUSTIFICACIONES, 1) = "1", True, False)
            MDIPrincipal.mnuRegNoMarc.Enabled = IIf(Mid(strAccesoMenu, MENU_REPORTE_NO_MARCACIONES, 1) = "1", True, False)
            MDIPrincipal.mnuRecordTard.Enabled = IIf(Mid(strAccesoMenu, MENU_REPORTE_TARDANZAS, 1) = "1", True, False)
            MDIPrincipal.mnuRecordDeInasist.Enabled = IIf(Mid(strAccesoMenu, MENU_REPORTE_INASISTENCIAS, 1) = "1", True, False)
            MDIPrincipal.mnuRecordDeSobretiempo.Enabled = IIf(Mid(strAccesoMenu, MENU_RECORD_SOBRETIEMPOS, 1) = "1", True, False)
            MDIPrincipal.mnuDetAsisTrab.Enabled = IIf(Mid(strAccesoMenu, MENU_DETALLE_ASISTENCIA_TRABAJADOR, 1) = "1", True, False)
            MDIPrincipal.mnuConsMens.Enabled = IIf(Mid(strAccesoMenu, MENU_CONSOLIDADO_MENSUAL_ASISTENCIA, 1) = "1", True, False)
            MDIPrincipal.mnuConsolidadoTrabajador.Enabled = IIf(Mid(strAccesoMenu, MENU_CONSOLIDADO_ASISTENCIA_TRABAJADOR, 1) = "1", True, False)
            MDIPrincipal.mnuBitacora.Enabled = IIf(Mid(strAccesoMenu, MENU_BITACORA, 1) = "1", True, False)
            
            Unload Me
            
            '////inicio
    With MMControl1
.Command = "stop" ' cerramos el archivo anterior
.Command = "close"
.FileName = App.Path & "\6.wav" ' y cargamos el nuevo
.Command = "open"
.Command = "play"
End With
    
    '///fin
            MDIPrincipal.Show
        Else
            MsgBox MSG_INVALIDATE_PASSWORD, vbInformation, mstrformname
            Me.txtClave.SetFocus
        End If
    End If
    
    Set objPersonal = Nothing
    Set rstMenuUsuario = Nothing
    Set rstUsuario = Nothing
    Screen.MousePointer = vbDefault
Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    Set objPersonal = Nothing
    Set rstUsuario = Nothing
    Set rstMenuUsuario = Nothing
    
    Call ErrorMessage(FrmSegAcceso_CmdAceptar_Click, Err.Source & " FrmSegAcceso : CmdAceptarClick", Err.Description)
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
    End
End Sub

Public Function ValidateFields() As Boolean
    ValidateFields = False
    If Trim$(Me.txtUsuario.Text) = "" Then
        MsgBox MSG_REQUERY_FIELD, vbInformation, mstrformname
        txtUsuario.SetFocus
        Exit Function
    End If
    Me.strUsuario = Me.txtUsuario.Text
    Me.strClave = Me.txtClave.Text
    ValidateFields = True
End Function

Private Sub Form_Activate()
'CmdAceptar_Click
End Sub

Private Sub Form_Load()
    gdatFecha_Nula_RangoInicio = CDate("01/01/1900")
    gdatFecha_Nula_RangoFin = CDate("31/12/2999")
    gdatHora_Nula_RangoInicio = CDate("00:00:00")
    gdatHora_Nula_RangoFin = CDate("23:59:59")
   
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub
