VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D95CB779-00CB-4B49-97B9-9F0B61CAB3C1}#4.0#0"; "biokey.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Begin VB.Form FrmConsTrabajador 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12135
   Icon            =   "FrmConsTrabajador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   12135
   Begin ZKFPEngXControl.ZKFPEngX ZKFPEngX2 
      Left            =   0
      Top             =   0
      EnrollCount     =   3
      SensorIndex     =   0
      Threshold       =   10
      VerTplFileName  =   ""
      RegTplFileName  =   ""
      OneToOneThreshold=   10
      Active          =   0   'False
      IsRegister      =   0   'False
      EnrollIndex     =   0
      SensorSN        =   ""
      FPEngineVersion =   "9"
      ImageWidth      =   0
      ImageHeight     =   0
      SensorCount     =   0
      TemplateLen     =   1152
      EngineValid     =   0   'False
      ForceSecondMatch=   0   'False
      IsReturnNoLic   =   -1  'True
      LowestQuality   =   30
   End
   Begin MCI.MMControl MMControl1 
      Height          =   735
      Left            =   1080
      TabIndex        =   4
      Top             =   6240
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1296
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdInit 
      Caption         =   "&Probar Huella"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   1335
   End
   Begin TrueOleDBGrid70.TDBGrid TdbgAdministrar 
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   10186
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
      Columns(3).Caption=   "Grupo"
      Columns(3).DataField=   "Grupo"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Tarjeta"
      Columns(4).DataField=   "Tarjeta"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Dirección"
      Columns(5).DataField=   "Direccion"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Telefono"
      Columns(6).DataField=   "Telefono"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "DNI"
      Columns(7).DataField=   "Dni"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=-39854098"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2752"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=71052096"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2752"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(2)._MinWidth=1"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=3281"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=2752"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=3281"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=3175"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(28)=   "Column(6).Width=2752"
      Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(31)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(32)=   "Column(7).Width=2752"
      Splits(0)._ColumnProps(33)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(35)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=94,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=91,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=92,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=93,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=98,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=95,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=96,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=97,.parent=17"
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
      Height          =   5625
      Left            =   11190
      TabIndex        =   0
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   9922
      BandCount       =   1
      FixedOrder      =   -1  'True
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   855
      _CBHeight       =   5625
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
         Height          =   570
         Left            =   30
         TabIndex        =   1
         Top             =   420
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
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Nuevo]"
               Object.ToolTipText     =   "Nuevo elemento"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Modificar]"
               Object.ToolTipText     =   "Modificar registro"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Eliminar]"
               Object.ToolTipText     =   "Eliminar registro"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Consultar]"
               Object.ToolTipText     =   "Pantalla de Consulta"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Imprimir]"
               Object.ToolTipText     =   "Imprimir registros"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Seleccionar]"
               Object.ToolTipText     =   "Seleccionar registro"
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImgLstAcciones 
      Left            =   10920
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsTrabajador.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsTrabajador.frx":64EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsTrabajador.frx":6DC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsTrabajador.frx":70E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsTrabajador.frx":CD02
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsTrabajador.frx":FA0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsTrabajador.frx":FD26
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsTrabajador.frx":FD84
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmConsTrabajador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strModo As enumActions
Public intIdTrabajador As Integer
Public strNombre As String
Public strApellido As String
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
Public datTime As Date
Private mrstDataSource As Object

Private Const mstrformname As String = "Listado de Trabajadores"

Dim FTempLenAdmin As Integer
Dim FingerCountAdmin As Long
Dim fpcHandleAdmin As Long
Dim FFingerHuellasAdmin() As String
Dim FFingerIdsAdmin() As Double
Dim FMatchTypeAdmin As Integer

Dim sTempAdmin As String
Dim FRegTemplateBKAdmin As Variant

Dim objPersonalHuellaAdmin As Object

Public Sub cmdVerify_Click()
    If ZKFPEngX2.IsRegister Then
        ZKFPEngX2.CancelEnroll
    End If
'    StatusBar.Caption = "Start Verification(1:1)"
    FMatchTypeAdmin = 1
End Sub


Public Property Get rstDataSource() As Variant
    Set rstDataSource = mrstDataSource
End Property
Public Property Set rstDataSource(ByVal vNewValue As Variant)
    Set mrstDataSource = vNewValue
    Set Me.TdbgAdministrar.DataSource = mrstDataSource
End Property

Private Sub cmdInit_Click()
  ZKFPEngX2.SensorIndex = 0
  If ZKFPEngX2.InitEngine = 0 Then
     'MessageBox 0, "Lector de Huellas iniciado correctamente", "Hint", 0
     
'     StatusBar.Caption = "Lector de Huellas iniciado correctamente"
'     TextSensorCount.Text = ZKFPEngX2.SensorCount & ""
'     TextSensorIndex.Text = ZKFPEngX2.SensorIndex & ""
'     TextSensorSN.Text = ZKFPEngX2.SensorSN
'
'     cmdInit.Enabled = False
     FMatchTypeAdmin = 0
     
     'ESTE ES EL VALOR A CAMBIAR, DEBEMOS AUMENTARLE o DISMINUIR PARA CONTROLAR REGULACION
     ZKFPEngX2.Threshold = 10
  End If
End Sub

Private Sub Form_Activate()
    Me.Tag = 1
    Colect.Add Me
End Sub

Private Sub Form_Load()
Dim objPersonal As Object
Dim rstTrabajador As Object
    Screen.MousePointer = vbHourglass
    FingerCountAdmin = 0
    fpcHandleAdmin = ZKFPEngX2.CreateFPCacheDB
    
    Set objPersonal = CreateObject("Personal.clsPersonal")
    'Me.datTime = Format(Time, "hh:mm")
    Set rstTrabajador = objPersonal.GetTrabajador()
    Set objPersonal = Nothing
    If rstTrabajador.RecordCount > 0 Then
        rstTrabajador.MoveFirst
        Do While Not rstTrabajador.EOF
            FRegTemplateBKAdmin = IIf(IsNull(rstTrabajador!Huella), "", rstTrabajador!Huella)
            
            'ZKFPEngX2.AddRegTemplateStrToFPCacheDB fpcHandleAdmin, FingerCountAdmin, FRegTemplateBKAdmin
                ReDim Preserve FFingerIdsAdmin(FingerCountAdmin + 1)
                ReDim Preserve FFingerHuellasAdmin(FingerCountAdmin + 1)
            FFingerIdsAdmin(FingerCountAdmin) = rstTrabajador!IdTrabajador
            FFingerHuellasAdmin(FingerCountAdmin) = FRegTemplateBKAdmin
            FingerCountAdmin = FingerCountAdmin + 1
            
            rstTrabajador.MoveNext
        Loop
        
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub ZKFPEngX2_OnCapture(ByVal ActionResult As Boolean, ByVal ATemplate As Variant)
    Dim fi As Long, i As Long
    Dim Score As Long, ProcessNum As Long
    Dim RegChanged As Boolean
    Dim FRegTemplateBKAdmin As String
    Dim verificaTrabajador As Boolean
    Dim objPersonal As Object
    Dim rstTrabajador As Object
    Screen.MousePointer = vbHourglass
    
    FRegTemplateBKAdmin = ""
    verificaTrabajador = False
    sTempAdmin = ZKFPEngX2.GetTemplateAsString()
  
    verificaTrabajador = False
    For i = 0 To FingerCountAdmin
        If ZKFPEngX2.VerFingerFromStr(FFingerHuellasAdmin(i), sTempAdmin, False, RegChanged) Then
            verificaTrabajador = True
            Exit For
       End If
    Next
    If verificaTrabajador Then
        'Me.datTime = Format(Time, "hh:mm")
        Me.intIdTrabajador = FFingerIdsAdmin(i)
        'If ValidateFieldsHuella Then
            
            Set objPersonal = New clsPersonal
            Set rstTrabajador = objPersonal.GetUnicoTrabajador(FFingerIdsAdmin(i))
            Set objPersonal = Nothing
            If rstTrabajador.RecordCount <= 0 Then
                Me.intIdTrabajador = 0
            Else
                Me.intIdTrabajador = rstTrabajador!IdTrabajador
            End If
    
            ' MUESTRA DATOS DEL TRABAJADOR
            'Me.Timer.Enabled = False
            'FrmMantDetalleMarcacion.intIdMarcacion = Me.intIdMarcacion
            
            FrmMantDetalleMarcacion.datTime = Me.datTime
            FrmMantDetalleMarcacion.TxtTarjeta.Text = rstTrabajador!Tarjeta
            FrmMantDetalleMarcacion.TxtApellido.Text = rstTrabajador!Apellido
            FrmMantDetalleMarcacion.txtNombre.Text = rstTrabajador!Nombre
            FrmMantDetalleMarcacion.strFoto = rstTrabajador!Foto
            'FrmMantDetalleMarcacion.strAccion = Me.strAccion
            FrmMantDetalleMarcacion.LblFecha = Date
            FrmMantDetalleMarcacion.graba = False
            FrmMantDetalleMarcacion.ShowForm
            
    '
    '
    '       FrmMantDetalleMarcacion.id = FFingerIdsAdmin(i)
    '       FrmMantDetalleMarcacion.intIdMarcacion = Me.intIdMarcacion
    '       FrmMantDetalleMarcacion.strAccion = Me.strAccion
    '       FrmMantDetalleMarcacion.ShowForm
        'End If
    Else
       FrmMensajeTrabajador.strMensaje = "Por favor vuelva a registrar su Huella"
       
        FrmMensajeTrabajador.Show
       
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub ZKFPEngX2_OnImageReceived(AImageValid As Boolean)
  ZKFPEngX2.PrintImageAt hDC, 8000, 5, ZKFPEngX2.ImageWidth, ZKFPEngX2.ImageHeight
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
On Error GoTo ErrHandler
    LoadData
    Me.Caption = mstrformname
    
    If Me.strModo = Find Then
        Me.Show vbModal 'OJO Estaba en Modal
        'Me.Show '**********taba asi********
    End If
    If Me.strModo = Query Then
        Me.Show
    End If
    
Exit Sub
ErrHandler:

    Call ErrorMessage(FrmConsTrabajador_ShowForm, Err.Source & " FrmConsTrabajador : ShowForm", Err.Description)
End Sub

Public Sub LoadData()
Dim objPersonal As Object

On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set rstDataSource = objPersonal.GetTrabajador(Me.intIdTrabajador, Me.strTarjeta, Me.strNombre, _
                                        Me.strApellido, Me.strDni, Me.intIdCategoria, _
                                        Me.intIdCargo, Me.intIdArea, Me.strFoto, _
                                        Me.strEstado, Me.datDesdeIngreso, Me.datHastaIngreso, _
                                        Me.datDesdeNacimiento, Me.datHastaNacimiento)
    Set objPersonal = Nothing
    Screen.MousePointer = vbDefault

Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    Set objPersonal = Nothing
    Call ErrorMessage(FrmConsTrabajador_LoadData, Err.Source & " FrmConsTrabajador : LoadData", Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Tag = 0
  Set mrstDataSource = Nothing
  ZKFPEngX2.FreeFPCacheDB (fpcHandleAdmin)
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
        Case KEY_ACTION_UPDATE
            Update
        Case KEY_ACTION_DELETE
            Delete
        Case KEY_ACTION_QUERY
            Finding
        Case KEY_ACTION_PRINT
            Imprimir
        Case KEY_ACTION_SELECCIONAR
            Encontrar
    End Select
End Sub

Public Function AddNew() As Boolean
Dim objform As FrmMantTrabajador

On Error GoTo ErrHandler
    If Me.strModo = Find Then
        MsgBox "La operación no está permitida...", vbExclamation, MSG_ACTION_ADDNEW & " - " & mstrformname
        AddNew = False
    Else
        Set objform = New FrmMantTrabajador
        objform.strAction = enumActions.AddNew
        objform.ShowForm
        Set objform = Nothing
        AddNew = True
    End If
Exit Function
ErrHandler:
    Set objform = Nothing
    AddNew = False
    Call ErrorMessage(FrmConsTrabajador_AddNew, Err.Source & " FrmConsTrabajador:AddNew", Err.Description)
End Function

Public Function Update() As Boolean
Dim objform As FrmMantTrabajador

On Error GoTo ErrHandler
    If Me.TdbgAdministrar.SelBookmarks.Count = 0 And Me.strModo = Find Then
        MsgBox MSG_SELECTION_EMPTY, vbExclamation, MSG_ACTION_UPDATE & " - " & mstrformname
        Update = False
    Else
        Set objform = New FrmMantTrabajador
        objform.strAction = enumActions.Update
        objform.intIdTrabajador = rstDataSource("IdTrabajador")
        objform.ShowForm
        Set objform = Nothing
        Update = True
    End If

Exit Function
ErrHandler:
    Set objform = Nothing
    Update = False
    Call ErrorMessage(FrmConsTrabajador_Update, Err.Source & " FrmConsTrabajador:Update", Err.Description)
End Function

Public Function Delete() As Boolean
Dim objPersonal As Object
Dim intIdTrabajador As Integer
Dim blnResult As Boolean
Dim strMensajeError As String
    
On Error GoTo ErrHandler
    If Me.TdbgAdministrar.SelBookmarks.Count = 0 And Me.strModo = Find Then
        MsgBox MSG_SELECTION_EMPTY, vbExclamation, MSG_ACTION_DELETE & " - " & mstrformname
        Delete = False
    Else
        If MsgBox(MSG_QUESTION_DELETE, vbOKCancel + vbQuestion, MSG_ACTION_DELETE & " - " & mstrformname) = 1 Then
            Screen.MousePointer = vbHourglass
            Set objPersonal = CreateObject("Personal.clsPersonal")
            intIdTrabajador = mrstDataSource("IdTrabajador")
            blnResult = objPersonal.DeleteTrabajador(intIdTrabajador, strMensajeError)
            If Not blnResult Then
                MsgBox strMensajeError, vbCritical, mstrformname
                Delete = False
            Else
                Bitacora "Trabajador", KEY_ACTION_DELETE, "Codigo " & Str(intIdTrabajador)
            End If
            Set objPersonal = Nothing
            Screen.MousePointer = vbDefault
            LoadData
        End If
    End If

Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Screen.MousePointer = vbDefault
    Delete = False
    Call ErrorMessage(FrmConsTrabajador_Delete, Err.Source & " FrmConsTrabajador:Delete", Err.Description)
End Function

Public Function Finding() As Boolean
    Unload Me
    FrmBusqTrabajador.strModo = Me.strModo
    If Not FrmBusqTrabajador.Visible Then
        FrmBusqTrabajador.ShowForm
    End If
End Function

Public Function Encontrar() As Boolean

    If Me.strModo = Find Then
        If Not TdbgAdministrar.FilterActive And TdbgAdministrar.Row >= 0 Then
            FrmBusqTrabajador.intIdTrabajador = rstDataSource("IdTrabajador")
            FrmBusqTrabajador.strApellidoNombre = rstDataSource("Apellido") & " " & rstDataSource("Nombre")
            FrmBusqTrabajador.strCadenaHorario = IIf(IsNull(rstDataSource("horario")), "", rstDataSource("horario"))
            'FrmBusqHorarioxTrabajador.intIdTrabajador = rstDataSource("IdTrabajador")
            'FrmBusqHorarioxTrabajador.strApellidoNombre = rstDataSource("Apellido") & " " & rstDataSource("Nombre")
            'FrmBusqHorarioxTrabajador.TxtIdTrabajador = rstDataSource("IdTrabajador")
            'FrmBusqHorarioxTrabajador.TxtApellido = rstDataSource("Apellido") & " " & rstDataSource("Nombre")
            Unload Me
            'FrmBusqHorarioxTrabajador.Show
            
        End If
    'Else
    'If Me.strModo = Seting Then
    '    FrmMantHorarioxTrabajador.TxtIdTrabajador = rstDataSource("IdTrabajador")
    '    FrmMantHorarioxTrabajador.TxtApellido = rstDataSource("Apellido") & " " & rstDataSource("Nombre")
    Else
        MsgBox "Error ... la acción no está disponible", vbCritical, mstrformname
    End If
    'End If
End Function
    
Public Function Imprimir() As Boolean
Dim rstTrabajadores As Object
Dim objTrabajadores As Object
Set objTrabajadores = CreateObject("Personal.clsReportes")
Set rstTrabajadores = objTrabajadores.GetReportTrabajadores(Me.intIdTrabajador, Me.strEstado, , Me.intIdCategoria, Me.intIdCargo, Me.intIdArea)
Set objTrabajadores = Nothing
ShowMultiReport rstTrabajadores, "rptRepListadoDeTrabajadores", gaParametrosReport()
Set rstTrabajadores = Nothing
End Function

