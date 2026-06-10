VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBusqTrabajador 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "FrmBusqTrabajador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7275
   Begin TabDlg.SSTab SstBusqueda 
      Height          =   3645
      Left            =   105
      TabIndex        =   18
      Top             =   120
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   6429
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Principal"
      TabPicture(0)   =   "FrmBusqTrabajador.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FraEstado"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Secundario"
      TabPicture(1)   =   "FrmBusqTrabajador.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FraIngreso"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FraNacimiento"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ChkFechaIngreso"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame2 
         Caption         =   "Datos Personales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   -74820
         TabIndex        =   30
         Top             =   405
         Width           =   5820
         Begin VB.TextBox TxtApellido 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1155
            MaxLength       =   60
            TabIndex        =   2
            Top             =   810
            Width           =   4545
         End
         Begin VB.TextBox TxtNombre 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1155
            MaxLength       =   60
            TabIndex        =   3
            Top             =   1290
            Width           =   4545
         End
         Begin VB.TextBox TxtIdTrabajador 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1155
            TabIndex        =   1
            Text            =   "0"
            Top             =   330
            Width           =   1095
         End
         Begin VB.TextBox TxtDni 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1155
            MaxLength       =   12
            TabIndex        =   4
            Top             =   1770
            Width           =   1815
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
            Left            =   195
            TabIndex        =   34
            Top             =   825
            Width           =   675
         End
         Begin VB.Label LblNombre 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
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
            Left            =   195
            TabIndex        =   33
            Top             =   1305
            Width           =   660
         End
         Begin VB.Label LblIdTrabajador 
            AutoSize        =   -1  'True
            Caption         =   "Código"
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
            Left            =   195
            TabIndex        =   32
            Top             =   345
            Width           =   570
         End
         Begin VB.Label LblDni 
            AutoSize        =   -1  'True
            Caption         =   "DNI"
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
            Left            =   195
            TabIndex        =   31
            Top             =   1785
            Width           =   300
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos de Planilla"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   195
         TabIndex        =   26
         Top             =   405
         Width           =   5730
         Begin MSDataListLib.DataCombo DtcCategoria 
            Height          =   315
            Left            =   1410
            TabIndex        =   8
            Top             =   225
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   714
            _Version        =   393216
            Style           =   2
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcCargo 
            Height          =   315
            Left            =   1410
            TabIndex        =   9
            Top             =   660
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   714
            _Version        =   393216
            Style           =   2
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcArea 
            Height          =   315
            Left            =   1410
            TabIndex        =   10
            Top             =   1110
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   714
            _Version        =   393216
            Style           =   2
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label LblCategoria 
            AutoSize        =   -1  'True
            Caption         =   "Categoria"
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
            Left            =   210
            TabIndex        =   29
            Top             =   255
            Width           =   825
         End
         Begin VB.Label LblCargo 
            AutoSize        =   -1  'True
            Caption         =   "Cargo"
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
            Left            =   210
            TabIndex        =   28
            Top             =   660
            Width           =   495
         End
         Begin VB.Label LblArea 
            AutoSize        =   -1  'True
            Caption         =   "Area"
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
            Left            =   210
            TabIndex        =   27
            Top             =   1140
            Width           =   405
         End
      End
      Begin VB.CheckBox ChkFechaIngreso 
         Caption         =   "Fecha de Ingreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   420
         TabIndex        =   11
         Top             =   1980
         Width           =   2085
      End
      Begin VB.Frame FraNacimiento 
         Caption         =   "           Fecha de Nacimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   195
         TabIndex        =   21
         Top             =   2715
         Width           =   5715
         Begin VB.CheckBox ChkFechaNacimiento 
            Caption         =   "Fecha de Nacimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   225
            TabIndex        =   14
            Top             =   0
            Width           =   2460
         End
         Begin MSComCtl2.DTPicker DtpDesdeNacimiento 
            Height          =   285
            Left            =   1365
            TabIndex        =   15
            Top             =   315
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   61341697
            CurrentDate     =   38036
         End
         Begin MSComCtl2.DTPicker DtpHastaNacimiento 
            Height          =   285
            Left            =   4065
            TabIndex        =   16
            Top             =   315
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   61341697
            CurrentDate     =   38036
         End
         Begin VB.Label LblHastaNacimiento 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Left            =   3135
            TabIndex        =   25
            Top             =   315
            Width           =   495
         End
         Begin VB.Label LblDesdeNacimiento 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            TabIndex        =   23
            Top             =   315
            Width           =   525
         End
      End
      Begin VB.Frame FraIngreso 
         Caption         =   "          Fecha de Ingreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   210
         TabIndex        =   20
         Top             =   2010
         Width           =   5715
         Begin MSComCtl2.DTPicker DtpDesdeIngreso 
            Height          =   285
            Left            =   1395
            TabIndex        =   12
            Top             =   240
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   61341697
            CurrentDate     =   38036
         End
         Begin MSComCtl2.DTPicker DtpHastaIngreso 
            Height          =   285
            Left            =   4050
            TabIndex        =   13
            Top             =   240
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   61341697
            CurrentDate     =   38036
         End
         Begin VB.Label LblHastaIngreso 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Left            =   3135
            TabIndex        =   24
            Top             =   285
            Width           =   495
         End
         Begin VB.Label LblDesdeIngreso 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Left            =   180
            TabIndex        =   22
            Top             =   285
            Width           =   525
         End
      End
      Begin VB.Frame FraEstado 
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74805
         TabIndex        =   19
         Top             =   2715
         Width           =   5790
         Begin VB.OptionButton OptEliminado 
            Caption         =   "Eliminado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton OptSuspendido 
            Caption         =   "Suspendido"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   6
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton OptNormal 
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   300
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin MSComctlLib.ImageList ImgLstAcciones 
      Left            =   6495
      Top             =   2580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBusqTrabajador.frx":0902
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBusqTrabajador.frx":360C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   2175
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   3836
      BandCount       =   1
      FixedOrder      =   -1  'True
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   855
      _CBHeight       =   2175
      _Version        =   "6.7.8988"
      BandBackColor1  =   -2147483635
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   795
      Width1          =   2955
      FixedBackground1=   0   'False
      UseCoolbarPicture1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   2340
         Left            =   30
         TabIndex        =   17
         Top             =   420
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   4128
         ButtonWidth     =   1376
         ButtonHeight    =   1376
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImgLstAcciones"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "               "
               Key             =   "[Consultar]"
               Object.ToolTipText     =   "Consultar datos"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Limpiar]"
               Object.ToolTipText     =   "Limpiar criterios de consulta"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmBusqTrabajador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strModo As enumActions
Public intIdTrabajador As Integer
Public strNombre As String
Public strApellido As String
Public strApellidoNombre As String
Public strDni As String
Public strFoto As String
Public strTarjeta As String
Public strEstado As String
Public datDesdeIngreso As Date
Public datDesdeNacimiento As Date
Public datHastaIngreso As Date
Public datHastaNacimiento As Date
Public intIdCategoria As Integer
Public intIdCargo As Integer
Public intIdArea As Integer
Public strCadenaHorario As String

Private mrstCategoria As Object
Private mrstCargo As Object
Private mrstArea As Object

Public blnStatus As Boolean
Public strMensajeError As String

Private Const mstrformname As String = "Consulta de Trabajadores"

Public Property Get rstCategoria() As Variant
    Set rstCategoria = mrstCategoria
End Property
Public Property Set rstCategoria(ByVal vNewValue As Variant)
    Set mrstCategoria = vNewValue
    If Not rstCategoria.EOF Then
        Set Me.DtcCategoria.RowSource = rstCategoria
        Me.DtcCategoria.ListField = "Descripcion"
        Me.DtcCategoria.BoundColumn = "IdCategoria"
        Me.DtcCategoria.Enabled = True
    Else
        Me.DtcCategoria.Enabled = False
        Me.DtcCategoria.Text = ""
    End If
End Property

Public Property Get rstCargo() As Variant
    Set rstCargo = mrstCargo
End Property
Public Property Set rstCargo(ByVal vNewValue As Variant)
    Set mrstCargo = vNewValue
    If Not rstCargo.EOF Then
        Set Me.DtcCargo.RowSource = rstCargo
        Me.DtcCargo.ListField = "Descripcion"
        Me.DtcCargo.BoundColumn = "IdCargo"
        Me.DtcCargo.Enabled = True
    Else
        Me.DtcCargo.Enabled = False
        Me.DtcCargo.Text = ""
    End If
End Property

Public Property Get rstArea() As Variant
    Set rstArea = mrstArea
End Property
Public Property Set rstArea(ByVal vNewValue As Variant)
    Set mrstArea = vNewValue
    If Not rstArea.EOF Then
        Set Me.DtcArea.RowSource = rstArea
        Me.DtcArea.ListField = "Descripcion"
        Me.DtcArea.BoundColumn = "IdArea"
        Me.DtcArea.Enabled = True
    Else
        Me.DtcArea.Enabled = False
        Me.DtcArea.Text = ""
    End If
End Property

Public Sub ShowForm()
On Error GoTo ErrHandler
    LoadData
    If Me.strModo = Find Then
        Me.Show vbModal
        'Me.Show   '******cambie aqui!!****
    End If
    If Me.strModo = Query Then
        Me.Show
    End If
    Me.ZOrder

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmBusqTrabajador_ShowForm, Err.Source & " FrmBusqTrabajador:ShowForm", Err.Description)
End Sub

Private Sub Form_Activate()
    Me.Tag = 1
    Colect.Add Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Me.Tag = 0
    Set mrstCategoria = Nothing
    Set mrstCargo = Nothing
    Set mrstArea = Nothing
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case KEY_ACTION_CLEAR
            Limpiar
        Case KEY_ACTION_QUERY
            Queries
    End Select
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.strMensajeError = ""
    Me.Caption = mstrformname
End Sub

Private Sub LoadData()
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set rstCategoria = objPersonal.GetCategoria()
    Set rstCargo = objPersonal.getcargo()
    Set rstArea = objPersonal.GetArea()
    Set objPersonal = Nothing
    
    Me.DtcCategoria.BoundText = Me.intIdCategoria
    Me.DtcCargo.BoundText = Me.intIdCargo
    Me.DtcArea.BoundText = Me.intIdArea
    Me.TxtIdTrabajador.Text = 0
    Me.TxtNombre.Text = ""
    Me.TxtApellido.Text = ""
    Me.TxtDni.Text = ""
    Me.OptNormal.Value = True
    Me.DtpDesdeIngreso.Value = Date
    Me.DtpDesdeNacimiento.Value = Date
    Me.DtpHastaIngreso.Value = Date
    Me.DtpHastaNacimiento.Value = Date
    Me.ChkFechaIngreso.Value = 0
    Me.ChkFechaNacimiento.Value = 0
    Me.SstBusqueda.Tab = 0
    
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Call ErrorMessage(FrmBusqTrabajador_LoadData, Err.Source & " FrmBusqTrabajador:LoadData", Err.Description)
End Sub

Public Function ValidateFields() As Boolean
    ValidateFields = False
    intIdTrabajador = Val(Me.TxtIdTrabajador.Text)
    strNombre = Replace(Me.TxtNombre.Text, "'", "''")
    strApellido = Replace(Me.TxtApellido.Text, "'", "''")
    strDni = Replace(Me.TxtDni.Text, "'", "''")
    If Me.ChkFechaIngreso.Value = 1 Then
        datDesdeIngreso = Me.DtpDesdeIngreso.Value
        datHastaIngreso = Me.DtpHastaIngreso.Value
    Else
        datDesdeIngreso = gdatFecha_Nula_RangoInicio
        datHastaIngreso = gdatFecha_Nula_RangoFin
    End If
    If Me.ChkFechaNacimiento.Value = 1 Then
        datDesdeNacimiento = Me.DtpDesdeNacimiento.Value
        datHastaNacimiento = Me.DtpHastaNacimiento.Value
    Else
        datDesdeNacimiento = gdatFecha_Nula_RangoInicio
        datHastaNacimiento = gdatFecha_Nula_RangoFin
    End If
    If Me.DtcCategoria.BoundText = "" Then
        intIdCategoria = 0
    Else
        intIdCategoria = Me.DtcCategoria.BoundText
    End If
    If Me.DtcCargo.BoundText = "" Then
        intIdCargo = 0
    Else
        intIdCargo = Me.DtcCargo.BoundText
    End If
    If Me.DtcArea.BoundText = "" Then
        intIdArea = 0
    Else
        intIdArea = Me.DtcArea.BoundText
    End If
    
    If Me.OptNormal.Value Then
        strEstado = gstrEstadoTrabajadorNormal
    End If
    If Me.OptEliminado.Value Then
        strEstado = gstrEstadoTrabajadorEliminado
    End If
    If Me.OptSuspendido.Value Then
        strEstado = gstrEstadoTrabajadorSuspendido
    End If
    ValidateFields = True
End Function

Private Function Queries() As Boolean

On Error GoTo ErrHandler:
    If Not ValidateFields Then
        Exit Function
    Else
        Screen.MousePointer = vbHourglass
        Queries = True
        FrmConsTrabajador.intIdTrabajador = Me.intIdTrabajador
        FrmConsTrabajador.strApellido = Me.strApellido
        FrmConsTrabajador.strNombre = Me.strNombre
        FrmConsTrabajador.strDni = Me.strDni
        FrmConsTrabajador.strEstado = Me.strEstado
        FrmConsTrabajador.strTarjeta = Me.strTarjeta
        FrmConsTrabajador.intIdCategoria = Me.intIdCategoria
        FrmConsTrabajador.intIdCargo = Me.intIdCargo
        FrmConsTrabajador.intIdArea = Me.intIdArea
        FrmConsTrabajador.datDesdeIngreso = Me.datDesdeIngreso
        FrmConsTrabajador.datDesdeNacimiento = Me.datDesdeNacimiento
        FrmConsTrabajador.datHastaIngreso = Me.datHastaIngreso
        FrmConsTrabajador.datHastaNacimiento = Me.datHastaNacimiento
        FrmConsTrabajador.strModo = Me.strModo
        FrmConsTrabajador.ShowForm
        Unload Me
        Screen.MousePointer = vbDefault
    End If

Exit Function
ErrHandler:
   Queries = False
   Screen.MousePointer = vbDefault

Call ErrorMessage(FrmBusqTrabajador_Queries, Err.Source & " FrmBusqTrabajador:Queries", Err.Description)
End Function

Private Function Limpiar() As Boolean
    Me.DtcCategoria.BoundText = Me.intIdCategoria
    Me.DtcCargo.BoundText = Me.intIdCargo
    Me.DtcArea.BoundText = Me.intIdArea
    Me.TxtIdTrabajador.Text = 0
    Me.TxtNombre.Text = ""
    Me.TxtApellido.Text = ""
    Me.TxtDni.Text = ""
    Me.OptNormal.Value = True
    Me.DtpDesdeIngreso.Value = Date
    Me.DtpDesdeNacimiento.Value = Date
    Me.DtpHastaIngreso.Value = Date
    Me.DtpHastaNacimiento.Value = Date
    Me.ChkFechaIngreso.Value = 1
    Me.ChkFechaNacimiento.Value = 0
    Me.SstBusqueda.Tab = 0
    Me.TxtApellido.SetFocus
    Limpiar = True
End Function


