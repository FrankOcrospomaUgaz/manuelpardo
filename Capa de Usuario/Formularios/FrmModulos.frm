VERSION 5.00
Begin VB.Form FrmModulos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   Icon            =   "FrmModulos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Módulos del Sistema"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdSeguridad 
         Height          =   1095
         Left            =   3480
         Picture         =   "FrmModulos.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Módulo de soporte de seguridad"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdPlanilla 
         Enabled         =   0   'False
         Height          =   1095
         Left            =   1800
         Picture         =   "FrmModulos.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Módulo para generación de planillas"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdPersonal 
         Height          =   1095
         Left            =   120
         Picture         =   "FrmModulos.frx":15D6
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Módulo para control de asistencia de personal"
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmModulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mstrformname As String = "Módulos del Sistema"


Private Sub cmdPersonal_Click()
    strSistema = Personal
    Me.Hide
    FrmSegAcceso.Show
End Sub

Private Sub cmdPlanilla_Click()
    strSistema = Planilla
    Me.Hide
    FrmSegAcceso.Show
End Sub

Private Sub cmdSeguridad_Click()
    strSistema = Seguridad
    Me.Hide
    FrmSegAcceso.Show
End Sub

Private Sub Form_Load()
    Me.Caption = mstrformname
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
