VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FrmMensajeTrabajador 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1845
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5865
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   2400
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "MSJ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1635
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5565
   End
End
Attribute VB_Name = "FrmMensajeTrabajador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Public strMensaje As String
Private Sub Form_Load()
i = 0
Me.Timer.Enabled = True
CenterForm Me
Me.lblMensaje.Caption = Me.strMensaje
 
End Sub

Private Sub Timer_Timer()
i = i + 1
If i = 2 Then
    Me.Timer.Enabled = False
    Unload Me
End If
End Sub
