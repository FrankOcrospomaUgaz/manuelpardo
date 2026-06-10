VERSION 5.00
Object = "{28D47522-CF84-11D1-834C-00A0249F0C28}#1.0#0"; "Gif89.dll"
Begin VB.Form FrmMensajeEspera 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1365
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin GIF89LibCtl.Gif89a Gif89a1 
      Height          =   195
      Left            =   1440
      OleObjectBlob   =   "FrmMensajeEspera.frx":0000
      TabIndex        =   1
      Top             =   960
      Width           =   2250
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "El sistema está procesando los datos. Esta acción puede demorar unos minutos...Espere un momento."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   465
      TabIndex        =   0
      Top             =   195
      Width           =   4095
   End
End
Attribute VB_Name = "FrmMensajeEspera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowForm()
    CenterForm Me
    Me.Show
End Sub
