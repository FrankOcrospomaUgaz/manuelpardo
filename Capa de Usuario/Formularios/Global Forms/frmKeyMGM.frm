VERSION 5.00
Begin VB.Form frmKeyMGM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CLAVE DE REGISTRO"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6915
   Icon            =   "frmKeyMGM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkBase 
      Caption         =   "Crear la Base de Datos"
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   2880
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Permite Crear la Base de Datos Personal en SQL SERVER 2000"
      Height          =   855
      Left            =   480
      TabIndex        =   14
      Top             =   2520
      Width           =   5655
   End
   Begin VB.TextBox txtd 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   5280
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   2070
      Width           =   705
   End
   Begin VB.TextBox txtc 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   4200
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   2070
      Width           =   705
   End
   Begin VB.TextBox txtb 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   3150
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2070
      Width           =   705
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3420
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4980
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtNombre 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2100
      TabIndex        =   0
      Top             =   1560
      Width           =   3945
   End
   Begin VB.TextBox txta 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   2100
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2070
      Width           =   705
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00CBB7A0&
      Height          =   315
      Left            =   5220
      Top             =   2010
      Width           =   855
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00CBB7A0&
      Height          =   315
      Left            =   4140
      Top             =   2010
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00CBB7A0&
      Height          =   315
      Left            =   3090
      Top             =   2010
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00CBB7A0&
      Height          =   315
      Left            =   2040
      Top             =   2010
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00CBB7A0&
      Height          =   315
      Left            =   2040
      Top             =   1500
      Width           =   4035
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5100
      TabIndex        =   8
      Top             =   1950
      Width           =   60
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4020
      TabIndex        =   7
      Top             =   1950
      Width           =   60
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2970
      TabIndex        =   6
      Top             =   1950
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A nombre de"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   1590
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clave de Registro"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   450
      TabIndex        =   4
      Top             =   2130
      Width           =   1260
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Registrar el Sistema"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000270B0&
      Height          =   255
      Left            =   1050
      TabIndex        =   13
      Top             =   90
      Width           =   3645
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Si Ud. cuenta con una clave de aurtorizacion por favor registre el sistema."
      ForeColor       =   &H004B4645&
      Height          =   330
      Left            =   1020
      TabIndex        =   12
      Top             =   420
      Width           =   5265
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   180
      Picture         =   "frmKeyMGM.frx":038A
      Top             =   60
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   -30
      Picture         =   "frmKeyMGM.frx":0958
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6990
   End
End
Attribute VB_Name = "frmKeyMGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdAceptar_Click()
Dim Cadena As String
Dim strClave As String
If chkBase.Value = 1 Then
    If MsgBox("Esta seguro de Crear la Base de Datos?" & Chr(13) & "Si existe la Base de Datos se borrara automaticamente", vbYesNo, "Creacion de la Base de Datos") = vbYes Then
           Crear_Base_Datos
    End If
End If
strClave = Me.txta & Me.txtb & Me.txtc & Me.txtd
If Trim(Me.txtNombre) = "" Then
    MsgBox "FALTA INGRESAR NOMBRE", vbExclamation, "MENSAJE"
ElseIf Trim(strClave) = "" Then
    MsgBox "FALTA INGRESAR CLAVE", vbExclamation, "MENSAJE"
ElseIf strClave <> GlobalMgm.Auth_ID Then
    MsgBox "LA CLAVE NO ES CORRECTA,INTENTE DE NUEVO", vbExclamation, "MENSAJE"
Else
SaveKeyForm App.Title, "Agree", "OK"
SaveKeyForm App.Title, "Enterprise", txtNombre
strClave = StrReverse(strClave)
'SaveKeyForm "System", "Identifier", strclave
APIMgm.GrabarClave "SOFTWARE\" & "Root" & "\" & "System", "Identifier", strClave, cHKEY_LOCAL_MACHINE
SaveKeyForm App.Title, "Identifier", strClave
MsgBox "SISTEMA REGISTRADO CORRECTAMENTE,GRACIAS!", vbInformation, "MENSAJE"
Unload Me
FrmPresentacion.Show
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub txta_Change()
If Len(Me.txta) = 4 Then Me.txtb.SetFocus
End Sub

Private Sub txtb_Change()
If Len(Me.txtb) = 4 Then Me.txtc.SetFocus
End Sub

Private Sub txtc_Change()
If Len(Me.txtc) = 4 Then Me.txtd.SetFocus
End Sub

Public Function Crear_Base_Datos() As Boolean
Dim objPersonal As Object
Dim Resultado As Boolean
Dim strSqlCreaBase As String
On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Seminario08.clsCreaBase")
    
    strSqlCreaBase = "IF EXISTS (SELECT * "
    strSqlCreaBase = strSqlCreaBase & "FROM   master..sysdatabases "
    strSqlCreaBase = strSqlCreaBase & "WHERE  name = 'Seminario08') "
    strSqlCreaBase = strSqlCreaBase & "DROP DATABASE Seminario08;"
    Resultado = objPersonal.Restaura(strSqlCreaBase)
'    If Resultado Then
'        End
'    End If
    
    strSqlCreaBase = "RESTORE DATABASE  Seminario08 " & Chr(13) & _
  "FROM DISK = '" & App.Path & "\BaseDatos\Seminario08.bak' " & Chr(13) & _
  "WITH FILE=1, " & Chr(13) & _
    "NORECOVERY;"
    Resultado = objPersonal.Restaura(strSqlCreaBase)
'    If Resultado Then
'        End
'    End If
    
    strSqlCreaBase = "RESTORE DATABASE Seminario08 WITH RECOVERY;"
    Resultado = objPersonal.Restaura(strSqlCreaBase)
'    If Resultado Then
'        End
'    End If
    
'    Set objPersonal = CreateObject("Personal.clsCrearBase")
'    Resultado = objPersonal.BorrarObjetos(ObtenerSentenciaBorraObjetos())
'    If Resultado Then
'        End
'    End If
'
'    Set objPersonal = CreateObject("Personal.clsCrearBase")
'    Resultado = objPersonal.BorrarObjetos(ObtenerSentenciaBorraObjetos())
'    If Resultado Then
'        End
'    End If
'
'    Set objPersonal = CreateObject("Personal.clsCrearBase")
'    Resultado = objPersonal.CrearTablas(ObtenerSentenciaCreaTablas())
'    If Resultado Then
'        End
'    End If
    Crear_Base_Datos = True
    Set objPersonal = Nothing
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Crear_Base_Datos = False

Call ErrorMessage(FrmMantMotivoJustificacion_AddNew, Err.Source & " FrmMantMotivoJustificacion:AddNew", Err.Description)

End Function
'
'Public Function ObtenerSentenciaBorraObjetos() As String
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Trabajador_Area]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[Trabajador] DROP CONSTRAINT FK_Trabajador_Area" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Trabajador_Cargo]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[Trabajador] DROP CONSTRAINT FK_Trabajador_Cargo" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Trabajador_Categoria]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[Trabajador] DROP CONSTRAINT FK_Trabajador_Categoria" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_DetalleConsolidadoAsistencia_ConsolidadoAsistencia]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[DetalleConsolidadoAsistencia] DROP CONSTRAINT FK_DetalleConsolidadoAsistencia_ConsolidadoAsistencia" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Trabajador_Grupo]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[Trabajador] DROP CONSTRAINT FK_Trabajador_Grupo" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Marcacion_HorarioxTrabajador]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[Marcacion] DROP CONSTRAINT FK_Marcacion_HorarioxTrabajador" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Marcacion_Justificacion]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[Marcacion] DROP CONSTRAINT FK_Marcacion_Justificacion" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Justificacion_MotivoJustificacion]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[Justificacion] DROP CONSTRAINT FK_Justificacion_MotivoJustificacion" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_AccesoUsuario_OpcionMenu]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[AccesoUsuario] DROP CONSTRAINT FK_AccesoUsuario_OpcionMenu" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_OpcionMenu_Sistema]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[OpcionMenu] DROP CONSTRAINT FK_OpcionMenu_Sistema" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_HorarioxTrabajador_TipoHora]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[HorarioxTrabajador] DROP CONSTRAINT FK_HorarioxTrabajador_TipoHora" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_DetalleConsolidadoAsistencia_Trabajador]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[DetalleConsolidadoAsistencia] DROP CONSTRAINT FK_DetalleConsolidadoAsistencia_Trabajador" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_HorarioxTrabajador_Trabajador]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[HorarioxTrabajador] DROP CONSTRAINT FK_HorarioxTrabajador_Trabajador" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Marcacion_Trabajador]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[Marcacion] DROP CONSTRAINT FK_Marcacion_Trabajador" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Suspension_Trabajador]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[Suspension] DROP CONSTRAINT FK_Suspension_Trabajador" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_AccesoUsuario_Usuario]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & Chr(13)
'strSql = strSql & "ALTER TABLE [dbo].[AccesoUsuario] DROP CONSTRAINT FK_AccesoUsuario_Usuario" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FormatDate]') and xtype in (N'FN', N'IF', N'TF'))" & Chr(13)
'strSql = strSql & "drop function [dbo].[FormatDate]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FormatDateCompleto]') and xtype in (N'FN', N'IF', N'TF'))" & Chr(13)
'strSql = strSql & "drop function [dbo].[FormatDateCompleto]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UP_TRANSACCION_MARCACION]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[UP_TRANSACCION_MARCACION]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarArea]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarArea]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarBitacora]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarBitacora]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarCargo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarCargo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarCategoria]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarCategoria]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarCliente]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarCliente]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarCompromiso]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarCompromiso]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarFeriado]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarFeriado]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarGrupo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarGrupo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarHorario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarHorario]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarHorarioxTrabajador]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarHorarioxTrabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarJustificacion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarJustificacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarJustificarFeriados]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarJustificarFeriados]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarMotivoJustificacion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarMotivoJustificacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarNuevoConsolidado]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarNuevoConsolidado]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarSuspension]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarSuspension]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarTipoHora]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarTipoHora]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarTrabajador]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarTrabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_AgregarUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_AgregarUsuario]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarAccesoUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarAccesoUsuario]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarArea]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarArea]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarCargo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarCargo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarCategoria]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarCategoria]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarCliente]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarCliente]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarDetalleConsolidadoAsistencia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarDetalleConsolidadoAsistencia]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarExisteConsolidado]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarExisteConsolidado]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarFaltas]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarFaltas]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarFaltasGrupo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarFaltasGrupo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarGrupo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarGrupo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarHolgura]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarHolgura]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarHorario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarHorario]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarHorarioxGrupo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarHorarioxGrupo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarHorarioxTrabajador]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarHorarioxTrabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarIncidYMinutoIns]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarIncidYMinutoIns]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarIncidYMinutoInsNo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarIncidYMinutoInsNo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarIncidYMinutoInsSinJust]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarIncidYMinutoInsSinJust]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarIncidYMinutoSalidaAnt]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarIncidYMinutoSalidaAnt]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarIncidYMinutoSalidaAntSinJust]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarIncidYMinutoSalidaAntSinJust]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarIncidYMinutoTard]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarIncidYMinutoTard]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarIncidYMinutoTardSinJust]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarIncidYMinutoTardSinJust]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarJornadasImcompletas]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarJornadasImcompletas]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarJustificacion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarJustificacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarMarcacion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarMarcacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarMarcacionHorarioxFecha]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarMarcacionHorarioxFecha]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarMarcacionModificarxFecha]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarMarcacionModificarxFecha]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarMarcacionesTrabajador]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarMarcacionesTrabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarMarcacionxFecha]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarMarcacionxFecha]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarMarcacionxTrabajador]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarMarcacionxTrabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarMesyAñoConsolidado]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarMesyAñoConsolidado]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarMinutoSuspensionTrabajador]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarMinutoSuspensionTrabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarMotivo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarMotivo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarMotivoJustificacion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarMotivoJustificacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarOpcionMenuDeUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarOpcionMenuDeUsuario]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarOpcionesModulo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarOpcionesModulo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarTardanzas]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarTardanzas]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarTolerancia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarTolerancia]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarTotalHoraDebioTrabajar]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarTotalHoraDebioTrabajar]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarTotalHoraExtra]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarTotalHoraExtra]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarTotalHoraTrabajada]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarTotalHoraTrabajada]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarTotalMinSuspension]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarTotalMinSuspension]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarTrabajador]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarTrabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarUnicoTrabajador]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarUnicoTrabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarUsuario]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_BuscarValidaUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_BuscarValidaUsuario]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_EliminarArea]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_EliminarArea]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_EliminarCargo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_EliminarCargo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_EliminarCategoria]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_EliminarCategoria]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_EliminarFeriado]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_EliminarFeriado]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_EliminarGrupo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_EliminarGrupo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_EliminarHorario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_EliminarHorario]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_EliminarHorarioxTrabajador]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_EliminarHorarioxTrabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_EliminarJustificacion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_EliminarJustificacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_EliminarMarcacion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_EliminarMarcacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_EliminarMotivoJustificacion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_EliminarMotivoJustificacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_EliminarSuspension]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_EliminarSuspension]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_EliminarTipoHora]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_EliminarTipoHora]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_EliminarTrabajador]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_EliminarTrabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_EliminarUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_EliminarUsuario]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarAccesoUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarAccesoUsuario]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarArea]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarArea]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarCargo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarCargo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarCategoria]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarCategoria]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarCliente]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarCliente]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarCodigoConsolidadoMarcacion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarCodigoConsolidadoMarcacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarConsolidadoAsistencia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarConsolidadoAsistencia]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarEstadoSuspendidoToNormal]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarEstadoSuspendidoToNormal]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarFeriado]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarFeriado]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarGrupo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarGrupo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarHolgura]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarHolgura]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarHorario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarHorario]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarHorarioxTrabajador]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarHorarioxTrabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarJustificacion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarJustificacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarMarcacion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarMarcacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarMarcacionIgualIdeal]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarMarcacionIgualIdeal]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarMarcacionManual]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarMarcacionManual]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarMarcacionPermiso]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarMarcacionPermiso]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarMotivoJustificacion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarMotivoJustificacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarSuspension]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarSuspension]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarTipoHora]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarTipoHora]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarTrabajador]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarTrabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarTrabajadorxCompromiso]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarTrabajadorxCompromiso]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ModificarUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ModificarUsuario]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteComparacionHoraTrab]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteComparacionHoraTrab]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteComparacionHoraTrabAsistencia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteComparacionHoraTrabAsistencia]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteComparacionHoraTrabL]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteComparacionHoraTrabL]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteConsolidados]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteConsolidados]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteDetalleAsistenciaxTrabajador]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteDetalleAsistenciaxTrabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteDetalleAsistenciaxTrabajadorJust]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteDetalleAsistenciaxTrabajadorJust]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteDetalleAsistenciaxTrabajadorL]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteDetalleAsistenciaxTrabajadorL]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteHorariosTrab]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteHorariosTrab]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteJustificacionTrab]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteJustificacionTrab]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteJustificacionTrabL]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteJustificacionTrabL]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRecordInasistenciasD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRecordInasistenciasD]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRecordInasistenciasL]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRecordInasistenciasL]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRecordInasistenciasSoloDiaD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRecordInasistenciasSoloDiaD]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRecordInasistenciasSoloDiaL]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRecordInasistenciasSoloDiaL]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRecordSobretiemposD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRecordSobretiemposD]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRecordSobretiemposL]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRecordSobretiemposL]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRecordTardanzasD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRecordTardanzasD]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRecordTardanzasL]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRecordTardanzasL]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRegistroMarcacionDia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRegistroMarcacionDia]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRegistroMarcacionDiaTipo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRegistroMarcacionDiaTipo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRegistroMarcacionFecha]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRegistroMarcacionFecha]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRegistroMarcacionFechaTipo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRegistroMarcacionFechaTipo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRegistroNoMarcacionesAmbasD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRegistroNoMarcacionesAmbasD]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRegistroNoMarcacionesAmbasL]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRegistroNoMarcacionesAmbasL]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRegistroNoMarcacionesD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRegistroNoMarcacionesD]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRegistroNoMarcacionesESD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRegistroNoMarcacionesESD]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRegistroNoMarcacionesESL]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRegistroNoMarcacionesESL]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteRegistroNoMarcacionesL]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteRegistroNoMarcacionesL]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteSuspensionesTrab]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteSuspensionesTrab]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[up_ReporteTrabajadores]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)" & Chr(13)
'strSql = strSql & "drop procedure [dbo].[up_ReporteTrabajadores]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AccesoUsuario]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[AccesoUsuario]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Area]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Area]" & Chr(13)
''''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Bitacora]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Bitacora]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Campos]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Campos]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cargo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Cargo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Categoria]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Categoria]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cliente]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Cliente]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ConsolidadoAsistencia]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[ConsolidadoAsistencia]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DetalleConsolidadoAsistencia]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[DetalleConsolidadoAsistencia]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Feriado]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Feriado]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Grupo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Grupo]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Holgura]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Holgura]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HorarioxTrabajador]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[HorarioxTrabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Justificacion]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Justificacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Marcacion]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Marcacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MotivoJustificacion]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[MotivoJustificacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[OpcionMenu]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[OpcionMenu]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Resultados]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Resultados]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SecuenciaCodificacion]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[SecuenciaCodificacion]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Sistema]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Sistema]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Suspension]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Suspension]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TipoHora]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[TipoHora]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Trabajador]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Trabajador]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Usuario]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & Chr(13)
'strSql = strSql & "drop table [dbo].[Usuario]" & Chr(13)
''strSql = strSql & "GO    " & Chr(13)
'
'ObtenerSentenciaBorraObjetos = strSql
'End Function
'
'Public Function ObtenerSentenciaCreaTablas() As String
'    strSql = strSql & "CREATE TABLE [dbo].[AccesoUsuario] (" & Chr(13)
'strSql = strSql & "   [IdUsuario] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [IdOpcion] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Habilitado] [bit] NOT NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
''''strSql = strSql & "GO" & Chr(13)
''
'strSql = strSql & "CREATE TABLE [dbo].[Area] (" & Chr(13)
'strSql = strSql & "   [IdArea] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Descripcion] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
''''strSql = strSql & "GO" & Chr(13)
''
'strSql = strSql & "CREATE TABLE [dbo].[Bitacora] (" & Chr(13)
'strSql = strSql & "   [Usuario] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Tabla] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Operacion] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Registro] [nvarchar] (100) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Fecha] [datetime] NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[Campos] (" & Chr(13)
'strSql = strSql & "   [FechaMediana] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [HoraLarga] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [Entero] [int] NULL ," & Chr(13)
'strSql = strSql & "   [Boolean] [bit] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Texto] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Doble] [float] NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[Cargo] (" & Chr(13)
'strSql = strSql & "   [IdCargo] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Descripcion] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[Categoria] (" & Chr(13)
'strSql = strSql & "   [IdCategoria] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Descripcion] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[Cliente] (" & Chr(13)
'strSql = strSql & "   [Cliente] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Tolerancia] [bit] NULL ," & Chr(13)
'strSql = strSql & "   [Tiempo] [int] NULL ," & Chr(13)
'strSql = strSql & "   [TiempoPrimeraMarcacion] [int] NULL ," & Chr(13)
'strSql = strSql & "   [HoraPrimeraMarcacion] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [ToleranciaTardanza] [int] NULL ," & Chr(13)
'strSql = strSql & "   [ToleranciaJornadaIncompleta] [int] NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[ConsolidadoAsistencia] (" & Chr(13)
'strSql = strSql & "   [CodigoConsolidado] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Mes] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Año] [nvarchar] (4) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [FechaInicio] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [FechaFin] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [TotalIncidenciaTardanza] [int] NULL ," & Chr(13)
'strSql = strSql & "   [TotalMinutoTardanza] [float] NULL ," & Chr(13)
'strSql = strSql & "   [TotalMinutoSuspension] [float] NULL ," & Chr(13)
'strSql = strSql & "   [TotalIncidenciaInasistencia] [int] NULL ," & Chr(13)
'strSql = strSql & "   [TotalMinutoInasistencia] [float] NULL ," & Chr(13)
'strSql = strSql & "   [TotalMinutoJornadaIncompleta] [float] NULL ," & Chr(13)
'strSql = strSql & "   [TotalHoraExtra] [float] NULL ," & Chr(13)
'strSql = strSql & "   [TotalHoraTrabajo] [float] NULL ," & Chr(13)
'strSql = strSql & "   [TotalHoraDebioTrabajar] [float] NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[DetalleConsolidadoAsistencia] (" & Chr(13)
'strSql = strSql & "   [CodigoConsolidado] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [IdTrabajador] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [IncidenciaTardanzaRemunerada] [int] NULL ," & Chr(13)
'strSql = strSql & "   [MinutoTardanzaRemunerado] [float] NULL ," & Chr(13)
'strSql = strSql & "   [IncidenciaTardanzaNoRemunerada] [int] NULL ," & Chr(13)
'strSql = strSql & "   [MinutoTardanzaNoRemunerado] [float] NULL ," & Chr(13)
'strSql = strSql & "   [IncidenciaInasistenciaRemunerada] [int] NULL ," & Chr(13)
'strSql = strSql & "   [MinutoInasistenciaRemunerada] [float] NULL ," & Chr(13)
'strSql = strSql & "   [IncidenciaFaltaInjustificada] [int] NULL ," & Chr(13)
'strSql = strSql & "   [MinutoFaltaInjustificada] [float] NULL ," & Chr(13)
'strSql = strSql & "   [IncidenciaSuspension] [int] NULL ," & Chr(13)
'strSql = strSql & "   [MinutoSuspension] [float] NULL ," & Chr(13)
'strSql = strSql & "   [HorasTrabajadas] [float] NULL ," & Chr(13)
'strSql = strSql & "   [IncidenciaJornadaIncompletaInjustificada] [int] NULL ," & Chr(13)
'strSql = strSql & "   [MinutoJornadaIncompletaInjustificada] [float] NULL ," & Chr(13)
'strSql = strSql & "   [IncidenciaFaltaJustificadaNoRemunerada] [int] NULL ," & Chr(13)
'strSql = strSql & "   [MinutoFaltaJustificadaNoRemunerada] [float] NULL ," & Chr(13)
'strSql = strSql & "   [IncidenciaJornadaIncompletaJustificadaRemunerada] [int] NULL ," & Chr(13)
'strSql = strSql & "   [MinutoJornadaIncompletaJustificadaRemunerada] [float] NULL ," & Chr(13)
'strSql = strSql & "   [IncidenciaJornadaIncompletaJustificadaNoRemunerada] [int] NULL ," & Chr(13)
'strSql = strSql & "   [MinutoJornadaIncompletaJustificadaNoRemunerada] [float] NULL ," & Chr(13)
'strSql = strSql & "   [IncidenciaFeriadoRemunerado] [int] NULL ," & Chr(13)
'strSql = strSql & "   [MinutoFeriadoRemunerado] [float] NULL ," & Chr(13)
'strSql = strSql & "   [IncidenciaFeriadoNoRemunerado] [int] NULL ," & Chr(13)
'strSql = strSql & "   [MinutoFeriadoNoRemunerado] [float] NULL ," & Chr(13)
'strSql = strSql & "   [Observacion] [varchar] (150) COLLATE Modern_Spanish_CI_AS NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[Feriado] (" & Chr(13)
'strSql = strSql & "   [IdFeriado] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [FechaFeriado] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [Remunerado] [bit] NOT NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[Grupo] (" & Chr(13)
'strSql = strSql & "   [IdGrupo] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Descripcion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[Holgura] (" & Chr(13)
'strSql = strSql & "   [DescripcionHolgura] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [ValorMin] [int] NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[HorarioxTrabajador] (" & Chr(13)
'strSql = strSql & "   [IdHorario] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [IdTipoHora] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [IdTrabajador] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Fecha] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [HoraIngreso] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [HoraSalida] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [TotalHoras] [float] NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[Justificacion] (" & Chr(13)
'strSql = strSql & "   [IdJustificacion] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [IdMotivoJustificacion] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Fecha] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [Remunerada] [nvarchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Observacion] [nvarchar] (200) COLLATE Modern_Spanish_CI_AS NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[Marcacion] (" & Chr(13)
'strSql = strSql & "   [IdMarcacion] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [IdJustificacion] [int] NULL ," & Chr(13)
'strSql = strSql & "   [IdTrabajador] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Fecha] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [HoraIngresoIdeal] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [HoraIngresoReal] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [HoraSalidaIdeal] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [HoraSalidaReal] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [IdHorario] [int] NULL ," & Chr(13)
'strSql = strSql & "   [CodigoConsolidado] [int] NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[MotivoJustificacion] (" & Chr(13)
'strSql = strSql & "   [IdMotivoJustificacion] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Descripcion] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[OpcionMenu] (" & Chr(13)
'strSql = strSql & "   [IdOpcion] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Descripcion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [PropiedadName] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [IdSistema] [int] NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[Resultados] (" & Chr(13)
'strSql = strSql & "   [IdMarcacion] [float] NULL ," & Chr(13)
'strSql = strSql & "   [IdJustificacion] [float] NULL ," & Chr(13)
'strSql = strSql & "   [IdTrabajador] [smallint] NULL ," & Chr(13)
'strSql = strSql & "   [Fecha] [smalldatetime] NULL ," & Chr(13)
'strSql = strSql & "   [HoraIngresoIdeal] [smalldatetime] NULL ," & Chr(13)
'strSql = strSql & "   [HoraIngresoReal] [smalldatetime] NULL ," & Chr(13)
'strSql = strSql & "   [HoraSalidaIdeal] [smalldatetime] NULL ," & Chr(13)
'strSql = strSql & "   [HoraSalidaReal] [smalldatetime] NULL ," & Chr(13)
'strSql = strSql & "   [IdHorario] [float] NULL ," & Chr(13)
'strSql = strSql & "   [CodigoConsolidado] [int] NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[SecuenciaCodificacion] (" & Chr(13)
'strSql = strSql & "   [Tabla] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Codigo] [float] NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[Sistema] (" & Chr(13)
'strSql = strSql & "   [IdSistema] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Descripcion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[Suspension] (" & Chr(13)
'strSql = strSql & "   [IdSuspension] [float] NOT NULL ," & Chr(13)
'strSql = strSql & "   [IdTrabajador] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Motivo] [nvarchar] (100) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [FechaInicio] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [FechaFin] [datetime] NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[TipoHora] (" & Chr(13)
'strSql = strSql & "   [IdTipoHora] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Descripcion] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Remunerada] [nvarchar] (1) COLLATE Modern_Spanish_CI_AS NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[Trabajador] (" & Chr(13)
'strSql = strSql & "   [IdTrabajador] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [IdGrupo] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [IdCategoria] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [IdArea] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [IdCargo] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [Tarjeta] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Apellido] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Nombre] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Dni] [nvarchar] (8) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Telefono] [nvarchar] (12) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Direccion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [FechaNacimiento] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [FechaIngreso] [datetime] NULL ," & Chr(13)
'strSql = strSql & "   [Foto] [nvarchar] (200) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Estado] [nvarchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [CompromisoGenerado] [bit] NOT NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'strSql = strSql & "CREATE TABLE [dbo].[Usuario] (" & Chr(13)
'strSql = strSql & "   [IdUsuario] [int] NOT NULL ," & Chr(13)
'strSql = strSql & "   [NombreUsuario] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ," & Chr(13)
'strSql = strSql & "   [Clave] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL " & Chr(13)
'strSql = strSql & ") ON [PRIMARY]" & Chr(13)
'''strSql = strSql & "GO" & Chr(13)
'
'ObtenerSentenciaCreaTablas = strSql
'End Function
'
