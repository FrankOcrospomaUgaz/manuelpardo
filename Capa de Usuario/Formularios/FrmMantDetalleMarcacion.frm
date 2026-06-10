VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FrmMantDetalleMarcacion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMantDetalleMarcacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7995
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   1200
      TabIndex        =   11
      Top             =   5520
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer 
      Left            =   165
      Top             =   1005
   End
   Begin VB.Frame FraTrabajador 
      Caption         =   "Datos Trabajador"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   4200
      Begin VB.TextBox TxtApellido 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   540
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1320
         Width           =   4005
      End
      Begin VB.TextBox TxtTarjeta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   360
         Left            =   945
         MaxLength       =   20
         TabIndex        =   3
         Top             =   315
         Width           =   3000
      End
      Begin VB.TextBox TxtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   555
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   2640
         Width           =   3975
      End
      Begin VB.Line Line2 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         DrawMode        =   1  'Blackness
         X1              =   45
         X2              =   4110
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label LblApellido 
         AutoSize        =   -1  'True
         Caption         =   "Apellido"
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
         Left            =   105
         TabIndex        =   8
         Top             =   945
         Width           =   675
      End
      Begin VB.Label LblTarjeta 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta"
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
         Left            =   105
         TabIndex        =   5
         Top             =   405
         Width           =   630
      End
      Begin VB.Label LbNombre 
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
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
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   750
      End
   End
   Begin VB.Frame FraFoto 
      Caption         =   "Foto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4605
      Left            =   4440
      TabIndex        =   0
      Top             =   930
      Width           =   3390
      Begin VB.Image ImgFoto 
         Height          =   4320
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3180
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "COLEGIO MANUEL PARDO                                             ""Saber Amar y Evangelizar"""
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   720
      TabIndex        =   7
      Top             =   960
      Width           =   3345
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      X1              =   600
      X2              =   4050
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label LblAccion 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   450
      Left            =   720
      TabIndex        =   6
      Top             =   1560
      Width           =   3345
   End
End
Attribute VB_Name = "FrmMantDetalleMarcacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intIdMarcacion As Long
Public strFoto As String
Public strAccion As String
Public datTime As Date
Public id As Double
Public graba As Boolean
Private Const mstrformname As String = "Detalle de Marcación"

Public Sub ShowForm()
On Error GoTo ErrHandler
Dim i As Integer
    LoadData
    Me.LblAccion.Caption = Me.strAccion
    Me.LblFecha.Caption = CStr(Date)
    Me.Show
    Me.ZOrder
    '////inicio
    With MMControl1
.Command = "stop" ' cerramos el archivo anterior
.Command = "close"
.FileName = App.Path & "\5.wav" ' y cargamos el nuevo
.Command = "open"
.Command = "play"
End With
    
    '///fin
    
    
Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantDetalleMarcacion_ShowForm, Err.Source & " FrmMantDetalleMarcacion:ShowForm", Err.Description)
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    graba = True
    CenterForm Me
End Sub

Private Sub LoadData()
Dim objPersonal As clsPersonal
Dim rst As Object
On Error Resume Next
'    Set objPersonal = New clsPersonal
'    Set rst = objPersonal.GetUnicoTrabajador(Me.id)
'
'    rst.MoveFirst
'
'    datTime = Me.datTime
'    TxtTarjeta.Text = rst!Tarjeta
'    TxtApellido.Text = rst!Apellido
'    txtNombre.Text = rst!Nombre
'    strFoto = rst!Foto
    LblFecha = Date
    Beep
    If strFoto <> "" Then
        Set Me.ImgFoto.Picture = LoadPicture(strFoto)
    Else
        Set Me.ImgFoto.Picture = LoadPicture("")
    End If
   
    Me.Timer.Enabled = True
    Me.Timer.Interval = 2000


Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantDetalleMarcacion_LoadData, Err.Source & " FrmMantDetalleMarcacion:LoadData", Err.Description)

End Sub

Private Sub Timer_Timer()
    If AddNew Then
        Me.Timer.Enabled = False
        FrmMantDetalleMarcacion.Hide
        If Me.graba Then
            FrmMantMarcacion.ShowForm
        End If
    End If
End Sub

Private Function AddNew() As Boolean
Dim objPersonal As clsPersonal

On Error GoTo ErrHandler
    If graba Then
        Set objPersonal = New clsPersonal
        AddNew = objPersonal.UpdateMarcacion(Me.intIdMarcacion, Me.datTime, Me.strAccion)
        Set objPersonal = Nothing
    Else
        AddNew = True
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    AddNew = False

    Call ErrorMessage(FrmMantDetalleMarcacion_AddNew, Err.Source & " FrmMantDetalleMarcacion:AddNew", Err.Description)
End Function
