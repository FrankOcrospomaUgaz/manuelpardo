VERSION 5.00
Begin VB.Form FrmMantHolguras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Holguras"
   ClientHeight    =   3060
   ClientLeft      =   6420
   ClientTop       =   3975
   ClientWidth     =   4605
   Icon            =   "FrmMantHolguras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4605
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
      Left            =   150
      TabIndex        =   5
      Top             =   1890
      Width           =   4125
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   2145
         Picture         =   "FrmMantHolguras.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cancelar la operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   1230
         Picture         =   "FrmMantHolguras.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
   End
   Begin VB.Frame fraHolguras 
      Caption         =   "Holgura"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   120
      TabIndex        =   0
      Top             =   630
      Width           =   4110
      Begin VB.TextBox txtSalida 
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
         Left            =   1245
         TabIndex        =   2
         Tag             =   "2"
         Top             =   735
         Width           =   1140
      End
      Begin VB.TextBox txtEntrada 
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
         Left            =   1245
         TabIndex        =   1
         Tag             =   "2"
         Top             =   285
         Width           =   1140
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
         Left            =   2685
         TabIndex        =   10
         Top             =   795
         Width           =   690
      End
      Begin VB.Label Label3 
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
         Left            =   2685
         TabIndex        =   9
         Top             =   345
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Salida"
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
         TabIndex        =   4
         Top             =   750
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Entrada"
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
         TabIndex        =   3
         Top             =   330
         Width           =   660
      End
   End
   Begin VB.Label LblComentario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Define la tolerancia de entrada y salida para considerar en el consolidado de asistencia generado por el sistema."
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
      Height          =   570
      Left            =   165
      TabIndex        =   8
      Top             =   60
      Width           =   4095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmMantHolguras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrformname As String = "Mantenimiento de Holguras"
Private Sub CmdAceptar_Click()
Dim objHolgura As Object
On Error GoTo ErrHandler
    Set objHolgura = CreateObject("Personal.clsPersonal")
    If objHolgura.UpdateHolgura(Int(Me.txtEntrada.Text), Int(Me.txtSalida.Text)) Then
        Bitacora "Holgura", KEY_ACTION_ADDNEW, "Entrada Nueva: " & Me.txtEntrada.Text & ", Salida Nueva: " & Me.txtSalida.Text
        MsgBox "Se Modificaron los datos correctamente", vbInformation, mstrformname
    Else
        MsgBox "No Se Modificaron los datos", vbCritical, "Error"
    End If
Me.txtEntrada.Text = objHolgura.GetHolgura("Entrada")
Me.txtSalida.Text = objHolgura.GetHolgura("Salida")
Exit Sub
ErrHandler:
   ErrorMessage FrmMantHolgura_CmdAceptar_Click, mstrformname & " - CmdAceptar_Click", Err.Description
End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Public Sub ShowForm()
On Error GoTo ErrHandler
Dim objHolgura As Object
Set objHolgura = CreateObject("Personal.clsPersonal")
Me.txtEntrada.Text = objHolgura.GetHolgura("Entrada")
Me.txtSalida.Text = objHolgura.GetHolgura("Salida")
Me.Show
Exit Sub
ErrHandler:
    ErrorMessage FrmMantHolgura_ShowForm, mstrformname & " - ShowForm", Err.Description
End Sub

Private Sub Form_Load()
    Me.Caption = mstrformname
    CenterForm Me
End Sub
