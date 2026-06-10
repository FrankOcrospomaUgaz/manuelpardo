VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmConsEliminarHorario 
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   Icon            =   "FrmConsEliminarHorario.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   9915
   Begin TrueOleDBGrid70.TDBGrid TdbgAdministrar 
      Height          =   4290
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   7567
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "IdHorario"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Trabajador"
      Columns(1).DataField=   "Nombres"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Fecha"
      Columns(2).DataField=   "Fecha"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "HoraIngreso"
      Columns(3).DataField=   "HoraIngreso"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "HoraSalida"
      Columns(4).DataField=   "HoraSalida"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1746"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1640"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=33"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4657"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4551"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=69974608"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2064"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1958"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(2)._MinWidth=69974608"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2196"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2090"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=4604"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=4498"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=15,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
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
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Named:id=33:Normal"
      _StyleDefs(53)  =   ":id=33,.parent=0"
      _StyleDefs(54)  =   "Named:id=34:Heading"
      _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(56)  =   ":id=34,.wraptext=-1"
      _StyleDefs(57)  =   "Named:id=35:Footing"
      _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   "Named:id=36:Selected"
      _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=37:Caption"
      _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(63)  =   "Named:id=38:HighlightRow"
      _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=39:EvenRow"
      _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(67)  =   "Named:id=40:OddRow"
      _StyleDefs(68)  =   ":id=40,.parent=33"
      _StyleDefs(69)  =   "Named:id=41:RecordSelector"
      _StyleDefs(70)  =   ":id=41,.parent=34"
      _StyleDefs(71)  =   "Named:id=42:FilterBar"
      _StyleDefs(72)  =   ":id=42,.parent=33"
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   3375
      Left            =   8985
      TabIndex        =   1
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   5953
      BandCount       =   1
      FixedOrder      =   -1  'True
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   855
      _CBHeight       =   3375
      _Version        =   "6.0.8169"
      BandBackColor1  =   -2147483635
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   795
      Width1          =   2955
      FixedBackground1=   0   'False
      UseCoolbarPicture1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   570
         Left            =   30
         TabIndex        =   2
         Top             =   345
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImgLstAcciones"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Eliminar]"
               Object.ToolTipText     =   "Eliminar registro"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Consultar]"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImgLstAcciones 
      Left            =   9000
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsEliminarHorario.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsEliminarHorario.frx":64EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsEliminarHorario.frx":6DC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsEliminarHorario.frx":70E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsEliminarHorario.frx":CD02
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmConsEliminarHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intIdGrupo As Integer
Public intIdTrabajador As Integer
Public datFechaInicio As Date
Public datFechaFin As Date
Public mrstDataSource As Object
Dim ArrayIdHorario() As String
Private Const mstrformname As String = "Consulta de Horarios A Eliminar"

Public Property Get rstDataSource() As Variant
    Set rstDataSource = mrstDataSource
End Property
Public Property Set rstDataSource(ByVal vNewValue As Variant)
    Set mrstDataSource = vNewValue
    Set Me.TdbgAdministrar.DataSource = mrstDataSource
End Property

Private Sub Form_Resize()
Dim intAuxTop As Integer
    
    On Error Resume Next
    intAuxTop = 100
    Me.TdbgAdministrar.Move 100, intAuxTop, (Me.Width - Me.ClbAcciones.Width - 300), (Me.Height - intAuxTop - 540)
    Me.ClbAcciones.Move Me.TdbgAdministrar.Width + 100, intAuxTop, Me.Width, (Me.Height - intAuxTop - 540)
End Sub

Public Sub ShowForm()
On Error GoTo ErrHandler
    LoadData
    Me.Caption = mstrformname
    Me.Show
Exit Sub
ErrHandler:
    Call ErrorMessage(FrmConsTrabajador_ShowForm, Err.Source & " FrmConsTrabajador : ShowForm", Err.Description)
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Caption = mstrformname
End Sub

Private Sub TDBGAdministrar_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If Not TdbgAdministrar.FilterActive And TdbgAdministrar.Row >= 0 Then
        TdbgAdministrar.SelBookmarks.Clear
        TdbgAdministrar.SelBookmarks.Add TdbgAdministrar.Bookmark
    End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_ACTION_DELETE
        If MsgBox("¿Desea Eliminar Todos los Horarios de La Lista?", vbQuestion + vbOKCancel, mstrformname) = vbOK Then
            Delete
        End If
    Case KEY_ACTION_QUERY
        Query
End Select
End Sub
Public Function Delete() As Boolean
On Error GoTo ErrHandler
    If Me.rstDataSource.RecordCount = 0 Then
        MsgBox MSG_EMPTY, vbInformation, mstrformname
        Exit Function
    End If
    If Not ConfiguraArray Then Exit Function
    If Not DeleteHorario Then MsgBox MSG_OPERATION_FAIL, vbCritical, mstrformname
    LoadData

Exit Function
ErrHandler:
    Call ErrorMessage(FrmConsEliminarHorario_Delete, Err.Source & " FrmConsEliminarHorario : Delete", Err.Description)
End Function

Private Function DeleteHorario() As Boolean
Dim objPersonal As Object
On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    Set objPersonal = CreateObject("Personal.clsHorario")
    DeleteHorario = objPersonal.DeleteHorarioYMarcacion(ArrayIdHorario)
    Set objPersonal = Nothing
    Screen.MousePointer = vbDefault

Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    Set objPersonal = Nothing
    Call ErrorMessage(FrmConsEliminarHorario_Delete, Err.Source & " FrmConsEliminarHorario : Delete", Err.Description)
End Function

Public Function Query() As Boolean
    Unload Me
    If Not FrmBusqEliminarHorario.Visible Then
        FrmBusqEliminarHorario.ShowForm
    End If
End Function
Public Sub LoadData()
Dim objPersonal As Object

On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    Set objPersonal = CreateObject("Personal.clsHorario")
    Set rstDataSource = objPersonal.GetHorarioGrupo(Me.datFechaInicio, Me.datFechaFin, Me.intIdGrupo, Me.intIdTrabajador)
    Set objPersonal = Nothing
    Screen.MousePointer = vbDefault

Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    Set objPersonal = Nothing
    Call ErrorMessage(FrmConsEliminarHorario_LoadData, Err.Source & " FrmConsEliminarHorario : LoadData", Err.Description)
End Sub

Public Function ConfiguraArray() As Boolean
Dim i As Integer
On Error GoTo ErrHandler
ConfiguraArray = False
i = 0
ReDim ArrayIdHorario(0 To Me.rstDataSource.RecordCount - 1)
    Me.rstDataSource.MoveFirst
    Do While Not Me.rstDataSource.EOF
        ArrayIdHorario(i) = Me.rstDataSource!IdHorario
        i = i + 1
        Me.rstDataSource.MoveNext
    Loop
    ConfiguraArray = True
Exit Function
ErrHandler:
    ConfiguraArray = False
    Call ErrorMessage(FrmConsEliminarHorario_ConfiguraArray, Err.Source & " FrmConsEliminarHorario : ConfiguraArray", Err.Description)
End Function
