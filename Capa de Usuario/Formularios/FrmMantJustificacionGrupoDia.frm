VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmMantJustificacionGrupoDia 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   Icon            =   "FrmMantJustificacionGrupoDia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   8880
   Begin VB.Frame Frame2 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   25
      Top             =   4320
      Width           =   3675
      Begin MSComCtl2.DTPicker DtpDia 
         Height          =   285
         Left            =   1080
         TabIndex        =   26
         Top             =   480
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   63438849
         CurrentDate     =   38056
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dia :"
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
         Left            =   480
         TabIndex        =   27
         Top             =   555
         Width           =   360
      End
   End
   Begin VB.CommandButton CmdAgregarFecha 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      Picture         =   "FrmMantJustificacionGrupoDia.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Búsqueda de Trabajador"
      Top             =   4440
      Width           =   585
   End
   Begin VB.CommandButton CmdQuitarFecha 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      Picture         =   "FrmMantJustificacionGrupoDia.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Búsqueda de Trabajador"
      Top             =   4920
      Width           =   585
   End
   Begin VB.Frame Frame1 
      Caption         =   "Trabajador Selecciondo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   4920
      TabIndex        =   21
      Top             =   240
      Width           =   3675
      Begin TrueOleDBGrid70.TDBGrid TDBGTrabajadorSelec 
         Height          =   3330
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   5874
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
         Columns(1).Caption=   "Nombres"
         Columns(1).DataField=   "Nombres"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1614"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1508"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2223"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2117"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(40)  =   "Named:id=33:Normal"
         _StyleDefs(41)  =   ":id=33,.parent=0"
         _StyleDefs(42)  =   "Named:id=34:Heading"
         _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(44)  =   ":id=34,.wraptext=-1"
         _StyleDefs(45)  =   "Named:id=35:Footing"
         _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=36:Selected"
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(49)  =   "Named:id=37:Caption"
         _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(51)  =   "Named:id=38:HighlightRow"
         _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=39:EvenRow"
         _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(55)  =   "Named:id=40:OddRow"
         _StyleDefs(56)  =   ":id=40,.parent=33"
         _StyleDefs(57)  =   "Named:id=41:RecordSelector"
         _StyleDefs(58)  =   ":id=41,.parent=34"
         _StyleDefs(59)  =   "Named:id=42:FilterBar"
         _StyleDefs(60)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.CommandButton CmdAgregar 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      Picture         =   "FrmMantJustificacionGrupoDia.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Búsqueda de Trabajador"
      Top             =   1320
      Width           =   585
   End
   Begin VB.CommandButton CmdAgregarVarios 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      Picture         =   "FrmMantJustificacionGrupoDia.frx":1590
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Búsqueda de Trabajador"
      Top             =   1800
      Width           =   585
   End
   Begin VB.CommandButton CmdQuitarVarios 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      Picture         =   "FrmMantJustificacionGrupoDia.frx":19D2
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Búsqueda de Trabajador"
      Top             =   2355
      Width           =   585
   End
   Begin VB.CommandButton CmdQuitar 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      Picture         =   "FrmMantJustificacionGrupoDia.frx":1E14
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Búsqueda de Trabajador"
      Top             =   2880
      Width           =   585
   End
   Begin VB.Frame FraTrabajador 
      Caption         =   "Trabajador"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   3795
      Begin TrueOleDBGrid70.TDBGrid TDBGTrabajador 
         Height          =   3330
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   5874
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
         Columns(1).Caption=   "Nombres"
         Columns(1).DataField=   "Nombres"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1614"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1508"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2223"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2117"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(40)  =   "Named:id=33:Normal"
         _StyleDefs(41)  =   ":id=33,.parent=0"
         _StyleDefs(42)  =   "Named:id=34:Heading"
         _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(44)  =   ":id=34,.wraptext=-1"
         _StyleDefs(45)  =   "Named:id=35:Footing"
         _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=36:Selected"
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(49)  =   "Named:id=37:Caption"
         _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(51)  =   "Named:id=38:HighlightRow"
         _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=39:EvenRow"
         _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(55)  =   "Named:id=40:OddRow"
         _StyleDefs(56)  =   ":id=40,.parent=33"
         _StyleDefs(57)  =   "Named:id=41:RecordSelector"
         _StyleDefs(58)  =   ":id=41,.parent=34"
         _StyleDefs(59)  =   "Named:id=42:FilterBar"
         _StyleDefs(60)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.Frame FraFecha 
      Caption         =   "Fechas Seleccionada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4920
      TabIndex        =   14
      Top             =   4320
      Width           =   3705
      Begin TrueOleDBGrid70.TDBGrid TDBGFechaSelec 
         Height          =   1125
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1984
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Fecha"
         Columns(0).DataField=   "Fecha"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(36)  =   "Named:id=33:Normal"
         _StyleDefs(37)  =   ":id=33,.parent=0"
         _StyleDefs(38)  =   "Named:id=34:Heading"
         _StyleDefs(39)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(40)  =   ":id=34,.wraptext=-1"
         _StyleDefs(41)  =   "Named:id=35:Footing"
         _StyleDefs(42)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(43)  =   "Named:id=36:Selected"
         _StyleDefs(44)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(45)  =   "Named:id=37:Caption"
         _StyleDefs(46)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(47)  =   "Named:id=38:HighlightRow"
         _StyleDefs(48)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(49)  =   "Named:id=39:EvenRow"
         _StyleDefs(50)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(51)  =   "Named:id=40:OddRow"
         _StyleDefs(52)  =   ":id=40,.parent=33"
         _StyleDefs(53)  =   "Named:id=41:RecordSelector"
         _StyleDefs(54)  =   ":id=41,.parent=34"
         _StyleDefs(55)  =   "Named:id=42:FilterBar"
         _StyleDefs(56)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.Frame FraDatos 
      Caption         =   "Datos de Justificación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   360
      TabIndex        =   3
      Top             =   5865
      Width           =   6780
      Begin VB.TextBox TxtIdJustificacion 
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
         Left            =   1230
         TabIndex        =   7
         Tag             =   "-1"
         Top             =   540
         Width           =   1560
      End
      Begin VB.TextBox TxtObservacion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1230
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "FrmMantJustificacionGrupoDia.frx":2256
         Top             =   1680
         Width           =   5370
      End
      Begin VB.OptionButton OptNoRemunerada 
         Caption         =   "No Remunerada"
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
         Left            =   3562
         TabIndex        =   5
         Top             =   210
         Width           =   1740
      End
      Begin VB.OptionButton OptRemunerada 
         Caption         =   "Remunerada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1552
         TabIndex        =   4
         Top             =   210
         Width           =   1395
      End
      Begin MSDataListLib.DataCombo DtcMotivoJustificacion 
         Height          =   315
         Left            =   1230
         TabIndex        =   8
         Top             =   915
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPFecha 
         Height          =   285
         Left            =   1230
         TabIndex        =   9
         Top             =   1305
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   63438849
         CurrentDate     =   38056
      End
      Begin VB.Label LblMotivo 
         AutoSize        =   -1  'True
         Caption         =   "Motivo"
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
         Left            =   135
         TabIndex        =   13
         Top             =   945
         Width           =   585
      End
      Begin VB.Label LblCodigo 
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
         Left            =   105
         TabIndex        =   12
         Top             =   555
         Width           =   570
      End
      Begin VB.Label LblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Left            =   135
         TabIndex        =   11
         Top             =   1350
         Width           =   495
      End
      Begin VB.Label LblObservacion 
         AutoSize        =   -1  'True
         Caption         =   "Observación"
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
         TabIndex        =   10
         Top             =   1680
         Width           =   1050
      End
   End
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
      Height          =   2325
      Left            =   7320
      TabIndex        =   0
      Top             =   5880
      Width           =   1335
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   240
         Picture         =   "FrmMantJustificacionGrupoDia.frx":225E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Confirmar operación"
         Top             =   480
         Width           =   810
      End
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   240
         Picture         =   "FrmMantJustificacionGrupoDia.frx":32A0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cancelar la operación"
         Top             =   1440
         Width           =   810
      End
   End
End
Attribute VB_Name = "FrmMantJustificacionGrupoDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public intIdJustificacion As Long
Public intIdMotivoJustificacion As Integer
Public datFecha As Date
Public strRemunerada As String
Public strObservacion As String
Private ArrayTrabajadorSelec() As Variant
Private ArrayFechaSelec() As Variant
Private intSeleccionados As Integer

Public blnStatus As Boolean
Public strMensajeError As String

Private mrstTrabajador As Object
Private mrstMotivoJustificacion As Object
Private mrstTrabajadorSelec As Object
Private mrstFecha As Object

Private Const mstrformname As String = "Mantenimiento de Justificaciones por Grupo Dias"

Public Property Set rstFecha(ByVal vNewValue As Variant)
    Set mrstFecha = vNewValue
End Property

Public Property Get rstFecha() As Variant
    Set rstFecha = mrstFecha
End Property

Public Property Set rstTrabajador(ByVal vNewValue As Variant)
    Set mrstTrabajador = vNewValue
End Property

Public Property Get rstTrabajador() As Variant
    Set rstTrabajador = mrstTrabajador
End Property

Public Property Set rstTrabajadorSelec(ByVal vNewValue As Variant)
    Set mrstTrabajadorSelec = vNewValue
End Property

Public Property Get rstTrabajadorSelec() As Variant
    Set rstTrabajadorSelec = mrstTrabajadorSelec
End Property

Public Property Get rstMotivoJustificacion() As Variant
    Set rstMotivoJustificacion = mrstMotivoJustificacion
End Property
Public Property Set rstMotivoJustificacion(ByVal vNewValue As Variant)
    Set mrstMotivoJustificacion = vNewValue
    If Not rstMotivoJustificacion.EOF Then
        Set Me.DtcMotivoJustificacion.RowSource = rstMotivoJustificacion
        Me.DtcMotivoJustificacion.ListField = "Descripcion"
        Me.DtcMotivoJustificacion.BoundColumn = "IdMotivoJustificacion"
        Me.DtcMotivoJustificacion.Enabled = True
    Else
        Me.DtcMotivoJustificacion.Enabled = False
        Me.DtcMotivoJustificacion.Text = ""
    End If
End Property

Public Sub ShowForm()
On Error GoTo ErrHandler
    Select Case intAction
        Case enumActions.AddNewGrupoDia
            LoadData
            SettingControls enumActions.AddNewGrupo, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_ADDNEWGROUPDIA
            Me.DtcMotivoJustificacion.BoundText = 0
            Me.DTPFecha.Value = Date
            Me.DtpDia = Date
            Me.TxtObservacion.Text = ""
            Me.TxtIdJustificacion.Text = 0
            Me.CmdQuitarVarios.Enabled = False
            Me.CmdQuitar.Enabled = False
            Me.CmdQuitarFecha.Enabled = 0
    End Select
    Me.Show
    Me.ZOrder 0
Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantJustificacionGrupoDia_ShowForm, Err.Source & " FrmMantJustificacionGrupoDia:ShowForm", Err.Description)
End Sub

Private Sub CmdAceptar_Click()
Dim blnResult As Boolean

On Error GoTo ErrHandler
    If Not ValidateFields Then Exit Sub
    blnResult = AddNew()
    If Not blnResult Then
        MsgBox MSG_OPERATION_FAIL & vbCrLf & Me.strMensajeError, vbCritical, mstrformname
        Me.blnStatus = False
    Else
        Unload Me
        If ActiveForm("FrmConsJustificacion") Then
            FrmConsJustificacion.LoadData
        End If
    End If

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantJustificacionGrupoDia_ShowForm, Err.Source & " FrmMantJustificacionGrupoDia:ShowForm", Err.Description)
End Sub

Private Sub CmdAgregar_Click()
    If rstTrabajador.EOF And rstTrabajador.BOF Then
        MsgBox MSG_EMPTY, vbExclamation, MSG_ACTION_ADDNEW & " - " & mstrformname
    Else
        If Me.TDBGTrabajador.SelBookmarks.Count = 0 Then
            MsgBox MSG_SELECTION_EMPTY, vbExclamation, MSG_ACTION_ADDNEW & " - " & mstrformname
        Else
            rstTrabajadorSelec.AddNew
            rstTrabajadorSelec!IdTrabajador = rstTrabajador!IdTrabajador
            rstTrabajadorSelec!Nombres = rstTrabajador!Nombres
            rstTrabajadorSelec.Update
            rstTrabajador.Delete
            rstTrabajador.Update
            Set Me.TDBGTrabajador.DataSource = rstTrabajador
            Set Me.TDBGTrabajadorSelec.DataSource = rstTrabajadorSelec
            Me.TDBGTrabajadorSelec.Columns(0).Caption = "Código"
            Me.TDBGTrabajadorSelec.Columns(0).AutoSize
            rstTrabajadorSelec.Sort = "Nombres"
        End If
    End If
    Me.CmdQuitarVarios.Enabled = 1
    Me.CmdQuitar.Enabled = 1
End Sub

Private Sub CmdAgregarFecha_Click()
    If rstFecha.RecordCount > 0 Then
        rstFecha.MoveFirst
        Do While Not rstFecha.EOF
            If rstFecha!Fecha = Me.DtpDia.Value Then
                Exit Sub
            End If
            rstFecha.MoveNext
        Loop
    End If
    Me.rstFecha.AddNew
    Me.rstFecha!Fecha = Me.DtpDia.Value
    Me.rstFecha.Update
    Set Me.TDBGFechaSelec.DataSource = rstFecha
    Me.CmdQuitarFecha.Enabled = 1
End Sub

Private Sub CmdAgregarVarios_Click()
Dim objPersonal As Object
Dim rstDetalle As Object
    If rstTrabajador.BOF And rstTrabajador.EOF Then
        MsgBox MSG_EMPTY, vbExclamation, MSG_ACTION_ADDNEW & " - " & mstrformname
    Else
        Set objPersonal = CreateObject("Personal.clsGeneral")
        Set rstTrabajador = objPersonal.GetCampos("Entero As IdTrabajador,Texto Nombres")
        Set rstTrabajadorSelec = objPersonal.GetCampos("Entero As IdTrabajador,Texto Nombres")
        Set Me.TDBGTrabajador.DataSource = Nothing
        Me.TDBGTrabajador.Columns(0).Caption = "Código"
        Me.TDBGTrabajador.Columns(0).AutoSize
        
        Set objPersonal = CreateObject("Personal.clsPersonal")
        Set rstDetalle = objPersonal.GetTrabajador()
        rstDetalle.Sort = "nombres"
        Set objPersonal = Nothing
        
        If rstDetalle.RecordCount > 0 Then
            rstDetalle.MoveFirst
            Do While Not rstDetalle.EOF
                rstTrabajadorSelec.AddNew
                rstTrabajadorSelec!IdTrabajador = rstDetalle!IdTrabajador
                rstTrabajadorSelec!Nombres = rstDetalle!Nombres
                rstTrabajadorSelec.Update
                rstDetalle.MoveNext
            Loop
            rstTrabajadorSelec.MoveFirst
            Set Me.TDBGTrabajadorSelec.DataSource = mrstTrabajadorSelec
        Else
            MsgBox "No existen incidencias registradas...", vbCritical, mstrformname
        End If
        
        Set Me.TDBGTrabajador.DataSource = Nothing
        Set mrstTrabajador = Nothing
        Set objPersonal = CreateObject("Personal.clsGeneral")
        Set rstTrabajador = objPersonal.GetCampos("Entero As IdTrabajador,Texto as Nombres")
        FrmMensajeEspera.Hide
    End If
    Me.CmdAgregarVarios.Enabled = False
    Me.CmdAgregar.Enabled = False
    Me.CmdQuitarVarios.Enabled = 1
    Me.CmdQuitar.Enabled = 1
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdQuitar_Click()
    If rstTrabajadorSelec.EOF And rstTrabajadorSelec.BOF Then
        MsgBox MSG_EMPTY, vbExclamation, MSG_ACTION_ADDNEW & " - " & mstrformname
    Else
        If Me.TDBGTrabajadorSelec.SelBookmarks.Count = 0 Then
            MsgBox MSG_SELECTION_EMPTY, vbExclamation, MSG_ACTION_ADDNEW & " - " & mstrformname
        Else
            rstTrabajador.AddNew
            rstTrabajador!IdTrabajador = rstTrabajadorSelec!IdTrabajador
            rstTrabajador!Nombres = rstTrabajadorSelec!Nombres
            rstTrabajador.Update
            rstTrabajadorSelec.Delete
            rstTrabajadorSelec.Update
            Set Me.TDBGTrabajador.DataSource = rstTrabajador
            Set Me.TDBGTrabajadorSelec.DataSource = rstTrabajadorSelec
            rstTrabajador.Sort = "Nombres"
        End If
    End If
    Me.CmdAgregarVarios.Enabled = 1
    Me.CmdAgregar.Enabled = 1
End Sub

Private Sub CmdQuitarFecha_Click()
    If rstFecha.RecordCount > 0 Then
        rstFecha.MoveFirst
        rstFecha.Delete
        rstFecha.Update
    End If
End Sub

Private Sub CmdQuitarVarios_Click()
Dim objPersonal As Object
Dim rstDetalle As Object

    If rstTrabajadorSelec.BOF And rstTrabajadorSelec.EOF Then
        MsgBox MSG_EMPTY, vbExclamation, MSG_ACTION_ADDNEW & " - " & mstrformname
    Else
        Set objPersonal = CreateObject("Personal.clsGeneral")
        Set rstTrabajador = objPersonal.GetCampos("Entero As IdTrabajador,Texto as Nombres")
        Set rstTrabajadorSelec = objPersonal.GetCampos("Entero As IdTrabajador,Texto Nombres")
        Set Me.TDBGTrabajadorSelec.DataSource = Nothing
        Me.TDBGTrabajadorSelec.Columns(0).Caption = "Código"
        Me.TDBGTrabajadorSelec.Columns(0).AutoSize
        
        Set objPersonal = CreateObject("Personal.clsPersonal")
        Set rstDetalle = objPersonal.GetTrabajador()
        rstDetalle.Sort = "nombres"
        Set objPersonal = Nothing
        
        If rstDetalle.RecordCount > 0 Then
            rstDetalle.MoveFirst
            Do While Not rstDetalle.EOF
                rstTrabajador.AddNew
                rstTrabajador!IdTrabajador = rstDetalle!IdTrabajador
                rstTrabajador!Nombres = rstDetalle!Nombres
                rstTrabajador.Update
                rstDetalle.MoveNext
            Loop
            rstTrabajador.MoveFirst
            Set Me.TDBGTrabajador.DataSource = mrstTrabajador
        Else
            MsgBox "No existen incidencias registradas...", vbCritical, mstrformname
        End If
    End If
    Me.CmdQuitarVarios.Enabled = False
    Me.CmdQuitar.Enabled = False
    Me.CmdAgregarVarios.Enabled = 1
    Me.CmdAgregar.Enabled = 1
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.strMensajeError = ""
End Sub

Private Sub LoadData()
Dim objPersonal As Object
Dim rstDetalle As Object

On Error GoTo ErrHandler
    
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set rstMotivoJustificacion = objPersonal.GetMotivoJustificacion()
    Set objPersonal = Nothing

    Set objPersonal = CreateObject("Personal.clsGeneral")
    Set rstTrabajador = objPersonal.GetCampos("Entero As IdTrabajador,Texto as Nombres")
        
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set rstDetalle = objPersonal.GetTrabajador()
    rstDetalle.Sort = "nombres"
    Set objPersonal = Nothing
    
    If rstDetalle.RecordCount > 0 Then
        rstDetalle.MoveFirst
        Do While Not rstDetalle.EOF
            rstTrabajador.AddNew
            rstTrabajador!IdTrabajador = rstDetalle!IdTrabajador
            rstTrabajador!Nombres = rstDetalle!Nombres
            rstTrabajador.Update
            rstDetalle.MoveNext
        Loop
        Set Me.TDBGTrabajador.DataSource = mrstTrabajador
        rstTrabajador.MoveFirst
        FrmMensajeEspera.Hide
    Else
        Me.TDBGTrabajador.MoveFirst
        Do While Not Me.TDBGTrabajador.EOF
            Me.TDBGTrabajador.Delete
            TDBGTrabajador.MoveFirst
        Loop
        FrmMensajeEspera.Hide
        MsgBox "No existen incidencias registradas...", vbCritical, mstrformname
    End If
    
    Set objPersonal = CreateObject("Personal.clsGeneral")
    Set rstTrabajadorSelec = objPersonal.GetCampos("Entero As IdTrabajador,Texto Nombres")
    Set objPersonal = Nothing
    
    Set objPersonal = CreateObject("Personal.clsGeneral")
    Set rstFecha = objPersonal.GetCampos("FechaMediana As Fecha")
    Set objPersonal = Nothing
Exit Sub
ErrHandler:
    Set objPersonal = Nothing

    Call ErrorMessage(FrmMantJustificacionGrupoDia_LoadData, Err.Source & " FrmMantJustificacionGrupoDia:LoadData", Err.Description)
End Sub

Private Function AddNew() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    
    Set objPersonal = CreateObject("Personal.clsMovimientos")
    AddNew = objPersonal.AddNewJustificacionGrupoDia(Me.datFecha, Me.intIdMotivoJustificacion, Me.strRemunerada, ArrayTrabajadorSelec, ArrayFechaSelec, Me.strObservacion)
    Set objPersonal = Nothing
    If AddNew Then
        Bitacora "Justificacion", KEY_ACTION_ADDNEW
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    AddNew = False

Call ErrorMessage(FrmMantJustificacionGrupoDia_AddNew, Err.Source & " FrmMantJustificacionGrupoDia:AddNew", Err.Description)
End Function

Private Function Update() As Boolean

End Function

Public Function ValidateFields() As Boolean
    ValidateFields = False
    If Trim$(Me.DtcMotivoJustificacion.Text) = "" Then
        MsgBox "Debe seleccionar el motivo de justificacion", vbInformation, mstrformname
        Exit Function
    End If
      
    If Me.rstTrabajadorSelec.RecordCount <= 0 Then
        MsgBox "Debe seleccionar trabajadores", vbInformation, mstrformname
        Exit Function
    End If
    rstTrabajadorSelec.MoveFirst

    If Not ValidateConsolidado Then
        MsgBox "Existe una Fecha Imposible de Justificar, es posible que se haya consolidado", vbInformation, mstrformname
        Exit Function
    End If
    
    If Not ConfiguraArray Then
        MsgBox "Debe seleccionar marcaciones para justificar", vbInformation, mstrformname
        Exit Function
    End If
    
    If Me.OptRemunerada.Value Then
        Me.strRemunerada = gstrHoraRemunerada
    Else
        Me.strRemunerada = gstrHoraNoRemunerada
    End If
    Me.intIdMotivoJustificacion = Me.DtcMotivoJustificacion.BoundText
    Me.strObservacion = Replace(Me.TxtObservacion.Text, "'", "''")
    Me.datFecha = Me.DTPFecha.Value
    ValidateFields = True

End Function

Public Function ConfiguraArray() As Boolean
Dim intContador As Integer
Dim intColumnas As Integer
Dim intCantidadColumnas As Integer
Dim intSelecionado As Integer
On Error GoTo ErrHandler
    '*****
    intCantidadColumnas = rstTrabajadorSelec.Fields.Count
    intSelecionado = rstTrabajadorSelec.RecordCount
    ReDim ArrayTrabajadorSelec(0 To intSelecionado - 1, 0 To intCantidadColumnas - 1)
    
    rstTrabajadorSelec.MoveFirst
    Do While Not rstTrabajadorSelec.EOF
        For intColumnas = 0 To intCantidadColumnas - 1
            ArrayTrabajadorSelec(intContador, intColumnas) = rstTrabajadorSelec.Fields(intColumnas).Value
        Next
        intContador = intContador + 1
        rstTrabajadorSelec.MoveNext
    Loop
    '*****
    intContador = 0
    intCantidadColumnas = rstFecha.Fields.Count
    intSelecionado = rstFecha.RecordCount
    ReDim ArrayFechaSelec(0 To intSelecionado - 1, 0 To intCantidadColumnas - 1)
    
    rstFecha.MoveFirst
    Do While Not rstFecha.EOF
        For intColumnas = 0 To intCantidadColumnas - 1
            ArrayFechaSelec(intContador, intColumnas) = rstFecha.Fields(intColumnas).Value
        Next
        intContador = intContador + 1
        rstFecha.MoveNext
    Loop
    
    ConfiguraArray = True
    
Exit Function
ErrHandler:
    ConfiguraArray = False

    Call ErrorMessage(FrmMantJustificacionGrupoDia_ConfiguraArray, Err.Source & " FrmMantJustificacionGrupoDia:ConfiguraArray", Err.Description)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mrstMotivoJustificacion = Nothing
    Set mrstTrabajador = Nothing
End Sub

Public Function ValidateConsolidado() As Boolean
Dim objConsolidado As Object
Dim blnResult As Boolean

On Error GoTo ErrHandler
    ValidateConsolidado = False
    Set objConsolidado = CreateObject("Personal.clsMovimientos")
    Me.rstFecha.MoveFirst
    Do While Not Me.rstFecha.EOF
        blnResult = objConsolidado.GetExisteConsolidado(Me.rstFecha!Fecha)
        If blnResult Then
            Set objConsolidado = Nothing
            Exit Function
        End If
    Me.rstFecha.MoveNext
    Loop
    Set objConsolidado = Nothing
    ValidateConsolidado = True
Exit Function
ErrHandler:
    Set objConsolidado = Nothing
    ValidateConsolidado = False
    Call ErrorMessage(FrmMantJustificacionGrupoDia_ValidateConsolidado, Err.Source & " FrmMantJustificacionGrupoDia:ValidateConsolidado", Err.Description)
End Function
