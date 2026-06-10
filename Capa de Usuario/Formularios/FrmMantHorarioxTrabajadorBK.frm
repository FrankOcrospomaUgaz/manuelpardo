VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmMantHorarioxTrabajadorBK 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   Icon            =   "FrmMantHorarioxTrabajadorBK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9015
   Begin TrueOleDBGrid70.TDBGrid TDBGSemana 
      Height          =   3015
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   5318
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Fecha"
      Columns(0).DataField=   "Fecha"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Ingreso"
      Columns(1).DataField=   "HoraIngreso"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Salida"
      Columns(2).DataField=   "HoraSalida"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2302"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2196"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2037"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1931"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(1)._MinWidth=7471205"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=1799"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1693"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(2)._MinWidth=7471205"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=56,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Named:id=33:Normal"
      _StyleDefs(49)  =   ":id=33,.parent=0"
      _StyleDefs(50)  =   "Named:id=34:Heading"
      _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   ":id=34,.wraptext=-1"
      _StyleDefs(53)  =   "Named:id=35:Footing"
      _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=36:Selected"
      _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=37:Caption"
      _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(59)  =   "Named:id=38:HighlightRow"
      _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=39:EvenRow"
      _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(63)  =   "Named:id=40:OddRow"
      _StyleDefs(64)  =   ":id=40,.parent=33"
      _StyleDefs(65)  =   "Named:id=41:RecordSelector"
      _StyleDefs(66)  =   ":id=41,.parent=34"
      _StyleDefs(67)  =   "Named:id=42:FilterBar"
      _StyleDefs(68)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGHorario 
      Height          =   5415
      Left            =   5160
      TabIndex        =   17
      Top             =   1200
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   9551
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Fecha"
      Columns(0).DataField=   "Fecha"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Ingreso"
      Columns(1).DataField=   "HoraIngreso"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Salida"
      Columns(2).DataField=   "HoraSalida"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2302"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2196"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2037"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1931"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(1)._MinWidth=7471205"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=1799"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1693"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(2)._MinWidth=7471205"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=56,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Named:id=33:Normal"
      _StyleDefs(49)  =   ":id=33,.parent=0"
      _StyleDefs(50)  =   "Named:id=34:Heading"
      _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   ":id=34,.wraptext=-1"
      _StyleDefs(53)  =   "Named:id=35:Footing"
      _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=36:Selected"
      _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=37:Caption"
      _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(59)  =   "Named:id=38:HighlightRow"
      _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=39:EvenRow"
      _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(63)  =   "Named:id=40:OddRow"
      _StyleDefs(64)  =   ":id=40,.parent=33"
      _StyleDefs(65)  =   "Named:id=41:RecordSelector"
      _StyleDefs(66)  =   ":id=41,.parent=34"
      _StyleDefs(67)  =   "Named:id=42:FilterBar"
      _StyleDefs(68)  =   ":id=42,.parent=33"
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
      Height          =   1035
      Left            =   90
      TabIndex        =   14
      Top             =   60
      Width           =   8790
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
         Left            =   1755
         TabIndex        =   1
         Tag             =   "-1"
         Text            =   "0"
         Top             =   195
         Width           =   975
      End
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
         Left            =   1740
         MaxLength       =   60
         TabIndex        =   2
         Tag             =   "-1"
         Top             =   585
         Width           =   6900
      End
      Begin VB.CommandButton CmdBuscarTrabajador 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8145
         TabIndex        =   3
         ToolTipText     =   "Búsqueda de Trabajador"
         Top             =   255
         Width           =   465
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
         Left            =   105
         TabIndex        =   16
         Top             =   270
         Width           =   570
      End
      Begin VB.Label LblApellido 
         AutoSize        =   -1  'True
         Caption         =   "Apellido y Nombre"
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
         TabIndex        =   15
         Top             =   660
         Width           =   1530
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
      Height          =   810
      Left            =   90
      TabIndex        =   13
      Top             =   5730
      Width           =   4875
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   2460
         Picture         =   "FrmMantHorarioxTrabajadorBK.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancelar la operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   1523
         Picture         =   "FrmMantHorarioxTrabajadorBK.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
   End
   Begin VB.Frame FraHorario 
      Caption         =   "Horario"
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
      Left            =   90
      TabIndex        =   0
      Top             =   1110
      Width           =   4920
      Begin VB.CommandButton CmdQuitarTodos 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4080
         Picture         =   "FrmMantHorarioxTrabajadorBK.frx":1C16
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Quitar todos los registros"
         Top             =   3360
         Width           =   705
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
         Height          =   330
         Left            =   3360
         Picture         =   "FrmMantHorarioxTrabajadorBK.frx":2058
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Agregar registros"
         Top             =   3360
         Width           =   705
      End
      Begin VB.Frame FraFecha 
         Caption         =   "Rango de Fechas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Width           =   4680
         Begin MSComCtl2.DTPicker DtpFechaFin 
            Height          =   330
            Left            =   3255
            TabIndex        =   5
            Top             =   285
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   582
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
            Format          =   20512769
            CurrentDate     =   38061
         End
         Begin MSComCtl2.DTPicker DtpFechaInicio 
            Height          =   330
            Left            =   870
            TabIndex        =   4
            Top             =   285
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   582
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
            Format          =   20512769
            CurrentDate     =   38061
         End
         Begin VB.Label LblFechaFin 
            AutoSize        =   -1  'True
            Caption         =   "Fin"
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
            Left            =   2640
            TabIndex        =   12
            Top             =   345
            Width           =   240
         End
         Begin VB.Label LblFechaInicio 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
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
            Left            =   150
            TabIndex        =   11
            Top             =   300
            Width           =   465
         End
      End
   End
End
Attribute VB_Name = "FrmMantHorarioxTrabajadorBK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public intIdHorario As Long
Public intIdTrabajador As Integer
Public intIdTipoHora As Integer
Public datFecha As Date
Public datHoraIngreso As Date
Public datHoraSalida As Date
Public blnStatus As Boolean
Public strMensajeError As String
Public tipo As Boolean

Public dFechaIni As Date
Public dFechaFin As Date

Private Array1() As String

Private ArrayTrabajador() As String
'Private mrstTipoHora As Object
Private mrstDetalle As Object
Private mrstSemana As Object

Private Const mstrformname As String = "Mantenimiento de Horario x Trabajador"

Public Property Set rstDetalle(ByVal vNewValue As Variant)
    Set mrstDetalle = vNewValue
End Property
Public Property Get rstDetalle() As Variant
    Set rstDetalle = mrstDetalle
End Property

Public Property Set rstSemana(ByVal vNewValue As Variant)
    If tipo Then
        Set mrstSemana = vNewValue
    Else
        Do While Not vNewValue.EOF
            mrstSemana.AddNew
            mrstSemana!IdTrabajador = Val(Me.TxtIdTrabajador.Text)
            mrstSemana!IdTipoHora = vNewValue!IdTipoHora
            mrstSemana!FECHA = vNewValue!FECHA
            mrstSemana!HoraIngreso = FormatDateTime(vNewValue!HoraIngreso, vbShortTime)
            mrstSemana!HoraSalida = FormatDateTime(vNewValue!HoraSalida, vbShortTime)
            mrstSemana!diaSemana = vNewValue!diaSemana
            mrstSemana.Update
            vNewValue.MoveNext
        Loop
    End If
End Property
Public Property Get rstSemana() As Variant
    Set rstSemana = mrstSemana
End Property


'Public Property Get rstTipoHora() As Variant
'    Set rstTipoHora = mrstTipoHora
'End Property
'Public Property Set rstTipoHora(ByVal vNewValue As Variant)
'    Set mrstTipoHora = vNewValue
'    If Not rstTipoHora.EOF Then
'        Set Me.dtcTipoHora.RowSource = rstTipoHora
'        Me.dtcTipoHora.ListField = "Descripcion"
'        Me.dtcTipoHora.BoundColumn = "IdTipoHora"
'        Me.dtcTipoHora.Enabled = True
'    Else
'        Me.dtcTipoHora.Enabled = False
'        Me.dtcTipoHora.Text = ""
'    End If
'End Property

Public Sub ShowForm()
On Error GoTo ErrHandler
    Select Case intAction
        Case enumActions.Update
            LoadData
            SettingControls enumActions.Update, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_UPDATE
        Case enumActions.AddNew
            LoadData
            SettingControls enumActions.AddNew, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_ADDNEW
            Me.TxtIdTrabajador.Text = 0
            Me.TxtApellido.Text = ""
            'Me.dtcTipoHora.BoundText = 0
            Me.DtpFechaInicio.Value = Date
            Me.DtpFechaFin.Value = Date
            'Me.DtpHoraIngreso.Value = Time
            'Me.DtpHoraSalida.Value = Time
            'Me.OptSabadoDomingo.Value = True
    End Select
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantHorarioxTrabajador_ShowForm, Err.Source & " FrmMantHorarioxTrabajador:ShowForm", Err.Description)
End Sub

Private Sub CmdAceptar_Click()
Dim blnResult As Boolean

On Error GoTo ErrHandler
    If Not ValidateFields Then Exit Sub
    Select Case intAction
        Case enumActions.Update
            blnResult = Update()
        Case enumActions.AddNew
            blnResult = AddNew()
    End Select

    If Not blnResult Then
        MsgBox MSG_OPERATION_FAIL & vbCrLf & Me.strMensajeError, vbCritical, mstrformname
        Me.blnStatus = False
    Else
        Unload Me
'        If ActiveForm("FrmConsHorarioxTrabajador") Then
'            FrmConsHorarioxTrabajador.LoadData
'        End If
    End If

Exit Sub
ErrHandler:

    Call ErrorMessage(FrmMantHorarioxTrabajador_ShowForm, Err.Source & " FrmMantHorarioxTrabajador:ShowForm", Err.Description)
End Sub

Private Sub CmdAgregar_Click()
Dim datFechaInicio As Date
Dim datFechaFin As Date
Dim diaSemana As Integer
    
    'rstDetalle.Filter = "IdTrabajador = " & Val(Me.TxtIdTrabajador.Text)
    
    datFechaInicio = Me.DtpFechaInicio.Value
    datFechaFin = Me.DtpFechaFin.Value
    
    
    If ValidateFieldsAgregar Then
        Do While datFechaInicio <= datFechaFin
            diaSemana = DateTime.WeekDay(datFechaInicio) - 1
            rstSemana.Filter = "DiaSemana = " & diaSemana
            Do While Not rstSemana.EOF
                rstDetalle.AddNew
                rstDetalle!IdTrabajador = Val(Me.TxtIdTrabajador.Text)
                rstDetalle!IdTipoHora = rstSemana!IdTipoHora
                rstDetalle!FECHA = datFechaInicio
                rstDetalle!HoraIngreso = FormatDateTime(rstSemana!HoraIngreso, vbShortTime)
                rstDetalle!HoraSalida = FormatDateTime(rstSemana!HoraSalida, vbShortTime)
                rstDetalle!diaSemana = rstSemana!diaSemana
                rstDetalle.Update
                rstSemana.MoveNext
            Loop
            datFechaInicio = datFechaInicio + 1
        Loop
    End If
    
            
    Set Me.TDBGHorario.DataSource = mrstDetalle
    Me.TDBGHorario.ReBind
        

End Sub

Private Sub CmdBuscarTrabajador_Click()
Dim objPersonal As Object
    FrmBusqTrabajador.strModo = Find
    FrmBusqTrabajador.ShowForm

    If Trim(FrmBusqTrabajador.intIdTrabajador) <> 0 Then
        Me.TxtIdTrabajador.Text = FrmBusqTrabajador.intIdTrabajador
        Me.TxtApellido.Text = Trim$(FrmBusqTrabajador.strApellidoNombre)
        Set objPersonal = CreateObject("Personal.clsPersonal")
        tipo = False
        Set rstSemana = objPersonal.GetHorarioSemana(FrmBusqTrabajador.intIdTrabajador, Date)
        Set objPersonal = Nothing
        tipo = True
        
        Set Me.TDBGSemana.DataSource = mrstSemana
        Me.TDBGSemana.ReBind
        
    End If
    Set FrmBusqTrabajador = Nothing
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdQuitar_Click()
Dim datFecha As Date

    If Me.TDBGHorario.SelBookmarks.Count = 0 Then
        MsgBox MSG_SELECTION_EMPTY, vbExclamation, MSG_ACTION_DELETE & " - " & mstrformname
    Else
        datFecha = rstDetalle!FECHA
        rstDetalle.Filter = "Fecha = " & datFecha
        rstDetalle.Delete
        rstDetalle.Filter = ""
        
        Set Me.TDBGHorario.DataSource = mrstDetalle
        Me.TDBGHorario.ReBind
    End If
End Sub

Private Sub CmdQuitarTodos_Click()
Dim objPersonal As Object

    Set objPersonal = CreateObject("Personal.clsGeneral")
    Set rstDetalle = objPersonal.GetCampos("Entero As IdTrabajador, FechaMediana as Fecha, HoraLarga As HoraIngreso, HoraLarga as HoraSalida")
    Set objPersonal = Nothing
    
    Set Me.TDBGHorario.DataSource = mrstDetalle
    Me.TDBGHorario.ReBind
    
Exit Sub
ErrHandler:
    Set objPersonal = Nothing

    Call ErrorMessage(FrmMantHorarioxTrabajador_LoadData, Err.Source & " FrmMantHorarioxTrabajador:LoadData", Err.Description)
End Sub

Private Sub Form_Load()
    CenterForm Me
    LoadData
    Me.strMensajeError = ""
End Sub

Private Sub LoadData()
Dim objPersonal As Object

On Error GoTo ErrHandler

    tipo = True
    Set objPersonal = CreateObject("Personal.clsGeneral")
    Set rstSemana = objPersonal.GetCampos("Entero As IdTrabajador, Entero as IdTipoHora, FechaMediana as Fecha, HoraLarga As HoraIngreso, HoraLarga as HoraSalida, Entero as DiaSemana")
    Set objPersonal = Nothing
    tipo = False
    Set objPersonal = CreateObject("Personal.clsGeneral")
    Set rstDetalle = objPersonal.GetCampos("Entero As IdTrabajador, Entero as IdTipoHora, FechaMediana as Fecha, HoraLarga As HoraIngreso, HoraLarga as HoraSalida, Entero as DiaSemana")
    Set objPersonal = Nothing
    
Exit Sub
ErrHandler:
    Set objPersonal = Nothing

    Call ErrorMessage(FrmMantHorarioxTrabajador_LoadData, Err.Source & " FrmMantHorarioxTrabajador:LoadData", Err.Description)
End Sub

Private Function AddNew() As Boolean
Dim objPersonal As Object
Screen.MousePointer = vbHourglass
On Error GoTo ErrHandler
    
    Set objPersonal = CreateObject("Personal.clsHorario")
    AddNew = objPersonal.AddNewHorarioxTrabajadorBK(Array1(), ArrayTrabajador())
    Set objPersonal = Nothing
    If AddNew Then
        Bitacora "Horario", KEY_ACTION_ADDNEW
    End If
    Screen.MousePointer = vbDefault
Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    Set objPersonal = Nothing
    AddNew = False

Call ErrorMessage(FrmMantHorarioxTrabajador_AddNew, Err.Source & " FrmMantHorarioxTrabajador:AddNew", Err.Description)
End Function
Private Function Update() As Boolean

End Function

Public Function ValidateFieldsAgregar() As Boolean
    ValidateFieldsAgregar = False
    
        
    If Me.DtpFechaFin.Value - Me.DtpFechaInicio.Value < 0 Then
        MsgBox MSG_INVALIDATE_DATE, vbInformation, mstrformname
        DtpFechaFin.SetFocus
        Exit Function
    End If
    ValidateFieldsAgregar = True
End Function

Public Function ValidateFields() As Boolean
    ValidateFields = False
    
    If Me.rstDetalle.RecordCount <= 0 Then
        MsgBox MSG_EMPTY, vbInformation, mstrformname
        Exit Function
    End If
    
    Me.intIdTrabajador = Val(Me.TxtIdTrabajador.Text)
    If ConfiguraArray Then
        ValidateFields = True
    End If
    
End Function

Private Function ConfiguraArray() As Boolean

Dim intContador As Integer
Dim dblTotalHoras As Double


On Error GoTo ErrHandler

intContador = 0
ReDim Array1(0 To rstDetalle.RecordCount)
' ************************ReDim Array2(0 To rstDetalle.RecordCount)
ReDim ArrayTrabajador(0 To 1)

rstDetalle.MoveFirst
Do While Not rstDetalle.EOF
    If intContador = 0 Then
        dFechaIni = rstDetalle!FECHA
    End If
    dblTotalHoras = DateDiff("n", rstDetalle!HoraIngreso, rstDetalle!HoraSalida)
    Array1(intContador) = "up_AgregarHorarioxTrabajador '" & rstDetalle!IdTipoHora & "','" & intIdTrabajador & "', '" & rstDetalle!FECHA & "','" & FormatDateTime(rstDetalle!HoraIngreso, vbShortTime) & "','" & FormatDateTime(rstDetalle!HoraSalida, vbShortTime) & "','" & dblTotalHoras & "'"
    'Array1(intContador) = "INSERT INTO HorarioxTrabajador (IdHorario, IdTipoHora, IdTrabajador, Fecha, HoraIngreso, HoraSalida, TotalHoras) VALUES ("
    'Array2(intContador) = ", " & Me.intIdTipoHora & ", " & Me.intIdTrabajador & ", '" & rstDetalle!Fecha & "','" & rstDetalle!HoraIngreso & _
                "','" & rstDetalle!HoraSalida & "', " & dblTotalHoras & ")"
    dFechaFin = rstDetalle!FECHA
    rstDetalle.MoveNext
    intContador = intContador + 1
Loop

ArrayTrabajador(0) = "up_ModificarTrabajadorxCompromisoBK '" & intIdTrabajador & "','0', '" & dFechaIni & "', '" & dFechaFin & "' "

ConfiguraArray = True

Exit Function
ErrHandler:
    ConfiguraArray = False

Call ErrorMessage(FrmMantHorarioxTrabajador_ConfiguraArray, Err.Source & " FrmMantHorarioxTrabajador:ConfiguraArray", Err.Description)



End Function

Private Sub Form_Unload(Cancel As Integer)
    'Set mrstTipoHora = Nothing
    Set mrstDetalle = Nothing
End Sub
