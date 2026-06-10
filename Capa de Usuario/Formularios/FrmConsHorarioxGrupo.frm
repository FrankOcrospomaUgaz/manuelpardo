VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmConsHorarioxGrupo 
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   Icon            =   "FrmConsHorarioxGrupo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   11400
   Begin TrueOleDBGrid70.TDBGrid TdbgAdministrar 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8705
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "IdTrabajador"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Apellidos"
      Columns(1).DataField=   "Apellido"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nombres"
      Columns(2).DataField=   "Nombre"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Fecha"
      Columns(3).DataField=   "Fecha"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "HoraIngreso"
      Columns(4).DataField=   "HoraIngreso"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "HoraSalida"
      Columns(5).DataField=   "HoraSalida"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "TipoHora"
      Columns(6).DataField=   "TipoHora"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Grupo"
      Columns(7).DataField=   "Grupo"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1217"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3334"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3228"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2963"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2858"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(2)._MinWidth=6648437"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=2117"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2011"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(4).Width=2090"
      Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=1984"
      Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(22)=   "Column(5).Width=1879"
      Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=1773"
      Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(26)=   "Column(6).Width=1693"
      Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=1588"
      Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(30)=   "Column(6)._MinWidth=-3"
      Splits(0)._ColumnProps(31)=   "Column(7).Width=2752"
      Splits(0)._ColumnProps(32)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(34)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(35)=   "Column(7)._MinWidth=-3"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=36,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=70,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=86,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=83,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=84,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=85,.parent=17"
      _StyleDefs(64)  =   "Named:id=33:Normal"
      _StyleDefs(65)  =   ":id=33,.parent=0"
      _StyleDefs(66)  =   "Named:id=34:Heading"
      _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   ":id=34,.wraptext=-1"
      _StyleDefs(69)  =   "Named:id=35:Footing"
      _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   "Named:id=36:Selected"
      _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=37:Caption"
      _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(75)  =   "Named:id=38:HighlightRow"
      _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=39:EvenRow"
      _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(79)  =   "Named:id=40:OddRow"
      _StyleDefs(80)  =   ":id=40,.parent=33"
      _StyleDefs(81)  =   "Named:id=41:RecordSelector"
      _StyleDefs(82)  =   ":id=41,.parent=34"
      _StyleDefs(83)  =   "Named:id=42:FilterBar"
      _StyleDefs(84)  =   ":id=42,.parent=33"
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   3735
      Left            =   10095
      TabIndex        =   1
      Top             =   60
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   6588
      BandCount       =   1
      FixedOrder      =   -1  'True
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   855
      _CBHeight       =   3735
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
         TabIndex        =   0
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
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Nuevo]"
               Object.ToolTipText     =   "Nuevo elemento"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Eliminar]"
               Object.ToolTipText     =   "Eliminar registro"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Consultar]"
               Object.ToolTipText     =   "Pantalla de consulta"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImgLstAcciones 
      Left            =   10230
      Top             =   4740
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
            Picture         =   "FrmConsHorarioxGrupo.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsHorarioxGrupo.frx":64EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsHorarioxGrupo.frx":6DC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsHorarioxGrupo.frx":70E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsHorarioxGrupo.frx":CD02
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmConsHorarioxGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strModo As enumActions
Public intIdGrupo As Integer
Public datFechaInicio As Date
Public datFechaFin As Date
Private mrstDataSource As Object

Private Const mstrformname As String = "Listado de Horarios x Grupo"

Public Property Get rstDataSource() As Variant
    Set rstDataSource = mrstDataSource
End Property
Public Property Set rstDataSource(ByVal vNewValue As Variant)
    Set mrstDataSource = vNewValue
    Set Me.TdbgAdministrar.DataSource = mrstDataSource
End Property

Private Sub Form_Load()
  LoadData
End Sub
Private Sub Form_Resize()
Dim intAuxTop As Integer
    
    On Error Resume Next
    intAuxTop = 100
    Me.TdbgAdministrar.Move 100, intAuxTop, (Me.Width - Me.ClbAcciones.Width - 300), (Me.Height - intAuxTop - 540)
    Me.ClbAcciones.Move Me.TdbgAdministrar.Width + 100, intAuxTop, Me.Width, (Me.Height - intAuxTop - 540)
End Sub
Public Sub CloseForm()
    genumActiveWindow = 0
    Unload Me
End Sub

Public Sub ShowForm()
    Me.Show
    CenterForm Me
    Me.Caption = mstrformname
End Sub

Public Sub LoadData()
Dim objPersonal As Object

On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    Set objPersonal = CreateObject("Personal.clsHorario")
    Set rstDataSource = objPersonal.GetHorarioxTrabajador(, , , , , Me.datFechaInicio, Me.datFechaFin, , , Me.intIdGrupo)
    Set objPersonal = Nothing
    Screen.MousePointer = vbDefault

Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    Set objPersonal = Nothing
    Call ErrorMessage(FrmConsHorarioxGrupo_LoadData, Err.Source & " FrmConsHorarioxGrupo : LoadData", Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set mrstDataSource = Nothing
End Sub

Private Sub TDBGAdministrar_HeadClick(ByVal ColIndex As Integer)
    OrdenaGrilla TdbgAdministrar.Columns(ColIndex).DataField, mrstDataSource, TdbgAdministrar
    Me.TdbgAdministrar.Col = ColIndex
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
        Case KEY_ACTION_ADDNEW
            AddNew
'        Case KEY_ACTION_UPDATE
'            Update
        Case KEY_ACTION_DELETE
            Delete
        Case KEY_ACTION_QUERY
            Finding
'        Case KEY_ACTION_PRINT
'            Imprimir
    End Select
End Sub

Public Function AddNew() As Boolean
Dim objform As FrmMantHorarioxGrupo

On Error GoTo ErrHandler
    Set objform = New FrmMantHorarioxGrupo
    objform.intAction = enumActions.AddNew
    objform.ShowForm
    Set objform = Nothing
    AddNew = True

Exit Function
ErrHandler:
    Set objform = Nothing
    AddNew = False
    Call ErrorMessage(FrmConsHorarioxGrupo_AddNew, Err.Source & " FrmConsHorarioxGrupo:AddNew", Err.Description)
End Function

Public Function Delete() As Boolean
Dim objPersonal As Object
Dim intIdHorario As Long
Dim blnResult As Boolean
Dim strMensajeError As String

On Error GoTo ErrHandler
    If Me.TdbgAdministrar.SelBookmarks.Count = 0 Then
        MsgBox MSG_SELECTION_EMPTY, vbExclamation, MSG_ACTION_DELETE & " - " & mstrformname
        Delete = False
    Else
        If MsgBox(MSG_QUESTION_DELETE, vbOKCancel + vbQuestion, MSG_ACTION_DELETE & " - " & mstrformname) = 1 Then
            Screen.MousePointer = vbHourglass
            intIdHorario = mrstDataSource("IdHorario")
            Set objPersonal = CreateObject("Personal.clsHorario")
            blnResult = objPersonal.DeleteHorarioxTrabajador(intIdHorario, strMensajeError)
            Set objPersonal = Nothing
            If Not blnResult Then
                MsgBox strMensajeError, vbCritical, mstrformname
                Delete = False
            Else
                Bitacora "HorarioxTrabajador", KEY_ACTION_DELETE, "Codigo " & Str(intIdHorario)
            End If
            Screen.MousePointer = vbDefault
            LoadData
        End If
    End If

Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Screen.MousePointer = vbDefault
    Delete = False
    Call ErrorMessage(FrmConsHorarioxGrupo_Delete, Err.Source & " FrmConsHorarioxGrupo:Delete", Err.Description)
End Function

Public Function Finding() As Boolean
    Unload Me
    FrmBusqHorarioxGrupo.strModo = Me.strModo
    FrmBusqHorarioxGrupo.ShowForm
End Function
