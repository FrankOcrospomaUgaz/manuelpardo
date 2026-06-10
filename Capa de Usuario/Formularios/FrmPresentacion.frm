VERSION 5.00
Begin VB.Form FrmPresentacion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   FillStyle       =   0  'Solid
   Icon            =   "FrmPresentacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtUsuario1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   360
      Top             =   240
   End
   Begin VB.Label TxtUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   6
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   1560
      X2              =   6240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   705
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6360
      Picture         =   "FrmPresentacion.frx":08CA
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Chiclayo - Perú"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Torres Paz N 337 - 2do Piso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   1560
      X2              =   6240
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPUSOFT SRL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de Control de Personal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   4380
   End
   Begin VB.Image ImgLogo 
      Height          =   900
      Left            =   5880
      Picture         =   "FrmPresentacion.frx":1194
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   3210
      Left            =   0
      Picture         =   "FrmPresentacion.frx":1B40
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   1575
   End
End
Attribute VB_Name = "FrmPresentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mstrformname As String = "Sistema de Control de Personal - Seminario Internacional"
Private Sub Form_Load()
    ConfiguraNombreCliente
    Me.TxtUsuario.Caption = gstrNameClient
    Me.Caption = mstrformname
End Sub

Private Sub Timer1_Timer()
    FrmPresentacion.Hide
    'FrmModulos.Show
    FrmSegAcceso.Show
    Timer1.Enabled = False
End Sub

