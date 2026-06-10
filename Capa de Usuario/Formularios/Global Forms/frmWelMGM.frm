VERSION 5.00
Begin VB.Form frmWelMGM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BIENVENIDO"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6855
   Icon            =   "frmWelMGM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4260
      TabIndex        =   1
      Top             =   3180
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Continuar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2220
      TabIndex        =   0
      Top             =   3180
      Width           =   1515
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Si Ud. cuenta con los componentes necesarios ya instalados haga click en Continuar."
      ForeColor       =   &H004B4645&
      Height          =   450
      Left            =   570
      TabIndex        =   7
      Top             =   2430
      Width           =   5985
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   120
      Picture         =   "frmWelMGM.frx":038A
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenido. Por favor siga los pasos especificados en este asistente."
      ForeColor       =   &H004B4645&
      Height          =   330
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   5265
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Intentando ejecutar por primera vez..."
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
      Left            =   1080
      TabIndex        =   5
      Top             =   90
      Width           =   4035
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Herramientas Cliente y/o Servidor de SQL Server version 2000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0095742F&
      Height          =   255
      Left            =   540
      TabIndex        =   4
      Top             =   1410
      Width           =   5655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2. Acesso a datos con OLEDB y Activex Data Object ADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0095742F&
      Height          =   195
      Left            =   540
      TabIndex        =   3
      Top             =   1710
      Width           =   4755
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Antes de continuar revisar si cuenta con los siguientes componentes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   90
      TabIndex        =   2
      Top             =   1020
      Width           =   5715
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   -30
      Picture         =   "frmWelMGM.frx":08C9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6960
   End
End
Attribute VB_Name = "frmWelMGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If APIMgm.LeerClave("SOFTWARE\" & "Root" & "\" & "System", "Identifier", cHKEY_LOCAL_MACHINE) = "" And APIMgm.getKeyForm(App.Title, "Identifier") = "" Then
APIMgm.SaveKeyForm App.Title, "Loaded", "OK"
APIMgm.SaveKeyForm App.Title, "Date", Date
Unload Me
frmKeyMGM.Show
Else
    FrmPruebaEvaluacion.Show
    FrmPruebaEvaluacion.TxtUsuario = FrmPresentacion.TxtUsuario & Chr(13) & "Faltan Archivos del Sistema, El Sistema ha sido modificado y no se ejecutará por Favor contactese con Compusoft"
End If
End Sub

