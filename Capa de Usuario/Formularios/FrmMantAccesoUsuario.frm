VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmMantAccesoUsuario 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "FrmMantAccesoUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   6735
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
      Left            =   120
      TabIndex        =   8
      Top             =   6555
      Width           =   6495
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   2494
         Picture         =   "FrmMantAccesoUsuario.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   3431
         Picture         =   "FrmMantAccesoUsuario.frx":190C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancelar la operación"
         Top             =   225
         Width           =   810
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGAdministrar 
      Height          =   5340
      Left            =   105
      TabIndex        =   1
      Top             =   1200
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   9419
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "IdOpcion"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   "Descripcion"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   84
      Columns(2)._MaxComboItems=   5
      Columns(2).ValueItems(0)._DefaultItem=   0
      Columns(2).ValueItems(0).Value=   "false"
      Columns(2).ValueItems(0).Value.vt=   8
      Columns(2).ValueItems(0).DisplayValue=   "FALSE"
      Columns(2).ValueItems(0).DisplayValue.vt=   8
      Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(1)._DefaultItem=   0
      Columns(2).ValueItems(1).Value=   "true"
      Columns(2).ValueItems(1).Value.vt=   8
      Columns(2).ValueItems(1).DisplayValue=   "TRUE"
      Columns(2).ValueItems(1).DisplayValue.vt=   8
      Columns(2).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems.Count=   2
      Columns(2).Caption=   "Habilitado"
      Columns(2).DataField=   "Habilitado"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1323"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1244"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=6509"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6429"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(1)._MinWidth=5"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(2)._MinWidth=40"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "Opciones del Menú"
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.locked=0"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   6510
      Begin VB.TextBox txtNombreUsuario 
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
         Left            =   1035
         TabIndex        =   7
         Tag             =   "-1"
         Top             =   675
         Width           =   5325
      End
      Begin VB.TextBox txtIdUsuario 
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
         Left            =   1035
         TabIndex        =   5
         Tag             =   "-1"
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label2 
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
         TabIndex        =   6
         Top             =   705
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código "
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
         TabIndex        =   4
         Top             =   315
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmMantAccesoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intIdUsuario As Integer
Public istrNombre As String
Public intAction As Integer
Public strNombreUsuario As String
Public mrstDataSource As Object

Private ArrayHabilitadas() As Variant

Private Const mstrformname As String = "Mantenimiento de Accesos de Usuario"

Public Property Get rstDataSource() As Variant
    Set rstDataSource = mrstDataSource
End Property
Public Property Set rstDataSource(ByVal vNewValue As Variant)
    Set mrstDataSource = vNewValue
    Set Me.TdbgAdministrar.DataSource = mrstDataSource
End Property

Private Sub CmdAceptar_Click()
On Error GoTo ErrHandler
Dim objPersonal As Object
Set objPersonal = CreateObject("Personal.clsSeguridad")

ConfiguraArray
If Not objPersonal.UpdateAccesoUsuario(Me.intIdUsuario, ArrayHabilitadas) Then
    MsgBox MSG_OPERATION_FAIL, vbCritical, mstrformname
Else
    MsgBox MSG_OPERATION_SUCESS, vbInformation, mstrformname
End If

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantAccesoUsuario_cmdAceptar, Err.Source & " FrmMantAccesoUsuario: cmdAceptar_Click", Err.Description)
End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
LoadData
CenterForm Me
End Sub
Public Sub LoadData()
Dim objPersonal As Object
On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsSeguridad")
    Set rstDataSource = objPersonal.GetOpcionMenuDeUsuario(Me.intIdUsuario)
    Set objPersonal = Nothing
    Exit Sub
ErrHandler:
        Call ErrorMessage(FrmMantAccesoUsuario_LoadData, Err.Source & " FrmMantAccesoUsuario: LoadData", Err.Description)
End Sub

Public Sub ShowForm()
Me.Show
Me.Caption = mstrformname
SettingControls enumActions.Update, Me
Me.txtIdUsuario.Text = Me.intIdUsuario
Me.txtNombreUsuario.Text = Me.strNombreUsuario
End Sub

Public Function ConfiguraArray() As Boolean
Dim intContador As Integer
Dim intColumnas As Integer
Dim intSeleccionados As Integer
Dim intCantidadColumnas As Integer

On Error GoTo ErrHandler

    
    intSeleccionados = rstDataSource.RecordCount
        
    intCantidadColumnas = rstDataSource.Fields.Count
    ReDim ArrayHabilitadas(0 To intSeleccionados - 1, 0 To intCantidadColumnas - 1)
    
    rstDataSource.MoveFirst
    Do While Not rstDataSource.EOF
            For intColumnas = 0 To intCantidadColumnas - 1
                If intColumnas = 2 Then
                    ArrayHabilitadas(intContador, intColumnas) = CInt(rstDataSource.Fields(intColumnas).Value) * -1
                Else
                    ArrayHabilitadas(intContador, intColumnas) = rstDataSource.Fields(intColumnas).Value
                End If
            Next
            intContador = intContador + 1
        rstDataSource.MoveNext
    Loop
    ConfiguraArray = True
    
Exit Function
ErrHandler:
    ConfiguraArray = False
    Call ErrorMessage(FrmMantAccesoUsuario_ConfiguraArray, Err.Source & " FrmMantAccesoUsuario: ConfiguraArray", Err.Description)
End Function

Private Function Update() As Boolean
Dim objPersonal As Object
On Error GoTo ErrHandler

    Set objPersonal = CreateObject("Personal.clsSeguridad")
    Update = objPersonal.UpdateAccesoUsuario(Me.intIdUsuario, ArrayHabilitadas)
    Set objPersonal = Nothing
    If Update Then
        Bitacora "AccesoUsuario", KEY_ACTION_UPDATE
    End If
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Update = False
Call ErrorMessage(FrmMantAccesoUsuario_Update, Err.Source & " FrmMantAccesoUsuario: Update", Err.Description)
End Function

