VERSION 5.00
Begin VB.Form frmMsgMGM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MENSAJE"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "frmMsgMGM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2010
      TabIndex        =   1
      Top             =   1740
      Width           =   1305
   End
   Begin VB.CheckBox chkShowAgain 
      Caption         =   "No mostrar este mensaje otra vez"
      Height          =   225
      Left            =   480
      TabIndex        =   0
      Top             =   1290
      Width           =   3075
   End
End
Attribute VB_Name = "frmMsgMGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
