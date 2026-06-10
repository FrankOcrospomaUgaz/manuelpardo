VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmMantCompromisoMarcacionUpdateGrupo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   Icon            =   "FrmMantCompromisoMarcacionUpdateGrupo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10650
   Begin VB.Frame Frame2 
      Caption         =   "Modo de modificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5880
      TabIndex        =   19
      Top             =   5880
      Width           =   4575
      Begin VB.CheckBox chkIgual 
         Caption         =   "Igualar hora REAL a hora IDEAL "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker DtpHoraSalidaReal 
         Height          =   285
         Left            =   1815
         TabIndex        =   12
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
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
         CustomFormat    =   "hh:mm tt"
         Format          =   50397187
         UpDown          =   -1  'True
         CurrentDate     =   38056
      End
      Begin MSComCtl2.DTPicker DtpHoraIngresoReal 
         Height          =   285
         Left            =   1815
         TabIndex        =   11
         Top             =   855
         Width           =   1335
         _ExtentX        =   2355
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
         CustomFormat    =   "hh:mm tt"
         Format          =   50397187
         UpDown          =   -1  'True
         CurrentDate     =   38056
      End
      Begin VB.Label LblHoraSalidaReal 
         AutoSize        =   -1  'True
         Caption         =   "Salida"
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
         TabIndex        =   21
         Top             =   1245
         Width           =   510
      End
      Begin VB.Label LblHoraIngresoReal 
         AutoSize        =   -1  'True
         Caption         =   " Ingreso"
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
         TabIndex        =   20
         Top             =   900
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   7440
      TabIndex        =   18
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton cmdDeseleccionar 
         Caption         =   "Quitar Selección "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   600
         TabIndex        =   4
         Top             =   480
         Width           =   2025
      End
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Seleccionar Todos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   600
         TabIndex        =   3
         Top             =   120
         Width           =   2025
      End
   End
   Begin VB.Frame FraFecha 
      Caption         =   "Seleccione el día para el cual modificar el registro de marcación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   7140
      Begin VB.CommandButton cmdConsultar 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6240
         Picture         =   "FrmMantCompromisoMarcacionUpdateGrupo.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Consultar inasistencias por fecha"
         Top             =   240
         Width           =   630
      End
      Begin MSComCtl2.DTPicker DtpFechaInicio 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   315
         Width           =   1785
         _ExtentX        =   3149
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
         Format          =   50397185
         CurrentDate     =   38056
      End
      Begin VB.Label LblFechaInicio 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
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
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   585
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
      Left            =   240
      TabIndex        =   15
      Top             =   6720
      Width           =   5580
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   2085
         Picture         =   "FrmMantCompromisoMarcacionUpdateGrupo.frx":35C4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Confirmar operación"
         Top             =   210
         Width           =   810
      End
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   3000
         Picture         =   "FrmMantCompromisoMarcacionUpdateGrupo.frx":4606
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cancelar la operación"
         Top             =   210
         Width           =   810
      End
   End
   Begin VB.Frame FraReal 
      Caption         =   "Horario a modificar"
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
      Left            =   240
      TabIndex        =   0
      Top             =   5880
      Width           =   5550
      Begin VB.OptionButton OptSoloSalida 
         Caption         =   "Sólo Salida"
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
         Left            =   2985
         TabIndex        =   8
         Top             =   300
         Width           =   1380
      End
      Begin VB.OptionButton OptSoloIngreso 
         Caption         =   "Sólo Ingreso"
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
         Left            =   1275
         TabIndex        =   7
         Top             =   300
         Width           =   1530
      End
      Begin VB.OptionButton OptNinguno 
         Caption         =   "Ninguno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   6
         Top             =   300
         Width           =   1110
      End
      Begin VB.OptionButton OptAmbas 
         Caption         =   "Ambas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4530
         TabIndex        =   9
         Top             =   300
         Width           =   960
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TdbgAdministrar 
      Height          =   4695
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   8281
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "IdMarcacion"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nombre Completo"
      Columns(1).DataField=   "NombreCompleto"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Ingreso Ideal"
      Columns(2).DataField=   "HoraIngresoIdeal"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Salida Ideal"
      Columns(3).DataField=   "HoraSalidaIdeal"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Ingreso Real"
      Columns(4).DataField=   "HoraIngresoReal"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Salida Real"
      Columns(5).DataField=   "HoraSalidaReal"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   84
      Columns(6)._MaxComboItems=   5
      Columns(6).ValueItems(0)._DefaultItem=   0
      Columns(6).ValueItems(0).Value=   "false"
      Columns(6).ValueItems(0).Value.vt=   8
      Columns(6).ValueItems(0).DisplayValue=   "FALSE"
      Columns(6).ValueItems(0).DisplayValue.vt=   8
      Columns(6).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(6).ValueItems(1)._DefaultItem=   0
      Columns(6).ValueItems(1).Value=   "true"
      Columns(6).ValueItems(1).Value.vt=   8
      Columns(6).ValueItems(1).DisplayValue=   "TRUE"
      Columns(6).ValueItems(1).DisplayValue.vt=   8
      Columns(6).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(6).ValueItems.Count=   2
      Columns(6).Caption=   "Modificar"
      Columns(6).DataField=   "Modificar"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1138"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=71144224"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4339"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=71153760"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2249"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2143"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=2143"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2037"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=2117"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2011"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(5).Width=2170"
      Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=2064"
      Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(27)=   "Column(6).Width=2752"
      Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(30)=   "Column(6).Order=7"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=216,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=70,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
      _StyleDefs(60)  =   "Named:id=33:Normal"
      _StyleDefs(61)  =   ":id=33,.parent=0"
      _StyleDefs(62)  =   "Named:id=34:Heading"
      _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   ":id=34,.wraptext=-1"
      _StyleDefs(65)  =   "Named:id=35:Footing"
      _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   "Named:id=36:Selected"
      _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=37:Caption"
      _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(71)  =   "Named:id=38:HighlightRow"
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "FrmMantCompromisoMarcacionUpdateGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public datFecha As Date
Private ArrayTrabajador() As Variant
Private intSeleccionados As Integer

'Public datHoraIngresoIdeal As Date
'Public datHoraSalidaIdeal As Date
Public datHoraIngresoReal As Date
Public datHoraSalidaReal As Date

Public blnStatus As Boolean
Public strMensajeError As String

Private mrstDetalle As Object
Private Const mstrformname As String = "Mantenimiento para Modificar Marcacion por Grupo"

Public Property Set rstDetalle(ByVal vNewValue As Variant)
    Set mrstDetalle = vNewValue
End Property
Public Property Get rstDetalle() As Variant
    Set rstDetalle = mrstDetalle
End Property

Public Sub ShowForm()
On Error GoTo ErrHandler
    Select Case intAction
        Case enumActions.UpdateGrupo
            LoadData
            SettingControls enumActions.Update, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_UPDATE
    End Select
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantCompromisoMarcacionDelete_ShowForm, Err.Source & " FrmMantCompromisoMarcacionDelete:ShowForm", Err.Description)
End Sub

Private Sub CmdAceptar_Click()
Dim blnResult As Boolean

On Error GoTo ErrHandler
    If Not ValidateFields Then Exit Sub
    Select Case intAction
        Case enumActions.UpdateGrupo
            blnResult = Update()
    End Select
    
    If Not blnResult Then
        MsgBox MSG_OPERATION_FAIL & vbCrLf & Me.strMensajeError, vbCritical, mstrformname
        Me.blnStatus = False
    Else
        Unload Me
        If ActiveForm("FrmConsCompromisoMarcacion") Then
            FrmConsCompromisoMarcacion.LoadData
        End If
    End If

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantCompromisoMarcacionDelete_ShowForm, Err.Source & " FrmMantCompromisoMarcacionDelete:ShowForm", Err.Description)
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdConsultar_Click()
Dim objPersonal As Object
Dim rstConsxEliminar As Object

On Error GoTo ErrHandler
        FrmMensajeEspera.ShowForm
        Set rstDetalle = Nothing
        Set objPersonal = CreateObject("Personal.clsGeneral")
        Set rstDetalle = objPersonal.GetCampos("Entero As IdMarcacion, Texto As NombreCompleto, HoraLarga As HoraIngresoIdeal, HoraLarga As HoraSalidaIdeal, HoraLarga As HoraIngresoReal, HoraLarga As HoraSalidaReal, Boolean as Modificar")
                
        Set objPersonal = CreateObject("Personal.clsMovimientos")
        Set rstConsxEliminar = objPersonal.GetMarcacionxModificar(Me.DtpFechaInicio)
        Set objPersonal = Nothing
        
        If rstConsxEliminar.RecordCount > 0 Then
            rstConsxEliminar.MoveFirst
            Do While Not rstConsxEliminar.EOF
                rstDetalle.AddNew
                rstDetalle!idmarcacion = rstConsxEliminar!idmarcacion
                rstDetalle!NombreCompleto = rstConsxEliminar!NombreCompleto
                rstDetalle!HoraIngresoIdeal = FormatDateTime(rstConsxEliminar!HoraIngresoIdeal, vbShortTime)
                rstDetalle!HoraSalidaIdeal = FormatDateTime(rstConsxEliminar!HoraSalidaIdeal, vbShortTime)
                If rstConsxEliminar!HoraIngresoReal <> "" Then
                    rstDetalle!HoraIngresoReal = FormatDateTime(rstConsxEliminar!HoraIngresoReal, vbShortTime)
                End If
                If rstConsxEliminar!HoraSalidaReal <> "" Then
                    rstDetalle!HoraSalidaReal = FormatDateTime(rstConsxEliminar!HoraSalidaReal, vbShortTime)
                End If
                rstDetalle!Modificar = False
                rstDetalle.Update
                rstConsxEliminar.MoveNext
            Loop
            Set Me.TdbgAdministrar.DataSource = mrstDetalle
            FrmMensajeEspera.Hide
        Else
            If rstConsxEliminar.RecordCount > 0 Then
            Me.TdbgAdministrar.MoveFirst
            Do While Not Me.TdbgAdministrar.EOF
                Me.TdbgAdministrar.Delete
                Me.TdbgAdministrar.MoveFirst
            Loop
            End If
            FrmMensajeEspera.Hide
            MsgBox "No existen incidencias registradas...", vbCritical, mstrformname
        End If
    
Exit Sub
ErrHandler:
    FrmMensajeEspera.Hide
    Set objPersonal = Nothing
    Set rstConsxEliminar = Nothing
    
    Call ErrorMessage(FrmMantCompromisoMarcacionDelete_LoadData, Err.Source & " FrmMantCompromisoMarcacionDelete:LoadData", Err.Description)
End Sub

Private Sub cmdDeseleccionar_Click()
    If rstDetalle.BOF And rstDetalle.EOF Then
        MsgBox "No existen trabajadores en la lista ", vbInformation, mstrformname
    Else
        rstDetalle.MoveFirst
        Do While Not rstDetalle.EOF
            If Me.rstDetalle!Modificar Then
                Me.rstDetalle!Modificar = False
            End If
            rstDetalle.Update
            rstDetalle.MoveNext
        Loop
    End If
End Sub

Private Sub cmdSeleccionar_Click()
    If rstDetalle.BOF And rstDetalle.EOF Then
        MsgBox "No existen trabajadores en la lista ", vbInformation, mstrformname
    Else
        rstDetalle.MoveFirst
        Do While Not rstDetalle.EOF
            If Not Me.rstDetalle!Modificar Then
                Me.rstDetalle!Modificar = True
            End If
            rstDetalle.Update
            rstDetalle.MoveNext
        Loop
        rstDetalle.MoveFirst
    End If
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.strMensajeError = ""
End Sub

Private Sub LoadData()
Dim objPersonal As Object

On Error GoTo ErrHandler

    Set objPersonal = CreateObject("Personal.clsGeneral")
    Set rstDetalle = objPersonal.GetCampos("Entero As IdMarcacion, Texto As NombreCompleto, HoraLarga As HoraIngresoIdeal, HoraLarga As HoraSalidaIdeal, HoraLarga As HoraIngresoReal, HoraLarga As HoraSalidaReal, Boolean as Modificar")
    Set objPersonal = Nothing
    Me.DtpFechaInicio.Value = Date
    
Exit Sub
ErrHandler:
    Set objPersonal = Nothing

    Call ErrorMessage(FrmMantCompromisoMarcacionDelete_LoadData, Err.Source & " FrmMantCompromisoMarcacionDelete:LoadData", Err.Description)
End Sub

Public Function ValidateFields() As Boolean
    ValidateFields = False
    If Me.rstDetalle.RecordCount <= 0 Then
        MsgBox "Debe seleccionar marcaciones", vbInformation, mstrformname
        Exit Function
    End If
    intSeleccionados = 0
    rstDetalle.MoveFirst
    Do While Not rstDetalle.EOF
        If rstDetalle!Modificar Then
            intSeleccionados = intSeleccionados + 1
        End If
        rstDetalle.MoveNext
    Loop
    If intSeleccionados = 0 Then
        MsgBox "Debe seleccionar marcaciones para Modificar", vbInformation, mstrformname
        Exit Function
    Else
        If ValidateConsolidado Then
            MsgBox "Existe una Fecha Imposible de Justificar, es posible que se haya consolidado", vbInformation, mstrformname
            Exit Function
        End If
        
        If Not ConfiguraArray Then
            MsgBox "Debe seleccionar marcaciones para justificar", vbInformation, mstrformname
            Exit Function
        End If
        If Not (Me.OptAmbas.Value Or Me.OptNinguno Or Me.OptSoloIngreso Or Me.OptSoloSalida) Then
            MsgBox "Debe seleccionar el Horario a Modificar", vbInformation, mstrformname
            Exit Function
        End If
        Me.datHoraIngresoReal = Me.DtpHoraIngresoReal.Value
        Me.datHoraSalidaReal = Me.DtpHoraSalidaReal.Value
        Me.datFecha = Me.DtpFechaInicio.Value
        
        ValidateFields = True
    End If
End Function

Public Function Update() As Boolean
Dim objPersonal As Object
Dim blnResult As Boolean
Dim strMensajeError As String
Dim enuOperacion As enumTipoModificacion

On Error GoTo ErrHandler
    If Me.TdbgAdministrar.SelBookmarks.Count < 0 Then
        MsgBox MSG_SELECTION_EMPTY, vbExclamation, MSG_ACTION_UPDATE & " - " & mstrformname
        Update = False
    Else
        If MsgBox(MSG_QUESTION_UPDATE, vbOKCancel + vbQuestion, MSG_ACTION_UPDATE & " - " & mstrformname) = 1 Then
            Screen.MousePointer = vbHourglass
            Set objPersonal = CreateObject("Personal.clsHorario")
            If Me.chkIgual.Value Then
                If Me.OptAmbas.Value Then
                    enuOperacion = Ambas
                ElseIf Me.OptSoloIngreso.Value Then
                        enuOperacion = SoloIngreso
                    ElseIf Me.OptSoloSalida.Value Then
                        enuOperacion = SoloSalida
                        ElseIf Me.OptNinguno.Value Then
                            enuOperacion = Ninguna
                End If
                Update = objPersonal.UpdateMarcacionGrupoIgualIdeal(ArrayTrabajador, enuOperacion)
            Else
                If Me.OptAmbas.Value Then
                    enuOperacion = Ambas
                    Update = objPersonal.UpdateMarcacionGrupoManual(ArrayTrabajador, , _
                                                Me.datHoraIngresoReal, , Me.datHoraSalidaReal, enuOperacion)
                ElseIf Me.OptSoloIngreso.Value Then
                        enuOperacion = SoloIngreso
                        Update = objPersonal.UpdateMarcacionGrupoManual(ArrayTrabajador, , _
                                                Me.datHoraIngresoReal, , , enuOperacion)
                    ElseIf Me.OptSoloSalida.Value Then
                        enuOperacion = SoloSalida
                        Update = objPersonal.UpdateMarcacionGrupoManual(ArrayTrabajador, , _
                                                 , , Me.datHoraSalidaReal, enuOperacion)
                        ElseIf Me.OptNinguno.Value Then
                            enuOperacion = Ninguna
                            Update = objPersonal.UpdateMarcacionGrupoManual(ArrayTrabajador, , _
                                                 , , , enuOperacion)
                End If
            End If
            Set objPersonal = Nothing
            
            If Update Then
                Bitacora "Marcacion", KEY_ACTION_UPDATE, "Grupo del dia " & CStr(Me.datFecha)
                Update = True
            End If
            Screen.MousePointer = vbDefault
        End If
    End If

Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Screen.MousePointer = vbDefault
    Update = False
    Call ErrorMessage(FrmConsCompromisoMarcacion_Delete, Err.Source & " FrmConsCompromisoMarcacion:Delete", Err.Description)
End Function

Private Function ConfiguraArray() As Boolean
Dim intIdMarcacion As Double
Dim objPersonal As Object

On Error GoTo ErrHandler

    ReDim ArrayTrabajador(0 To intSeleccionados)
    intSeleccionados = 0
    If Not rstDetalle.BOF And Not rstDetalle.EOF Then
        MsgBox "No existen marcaciones en la lista ", vbInformation, mstrformname
    Else
        rstDetalle.MoveFirst
        Do While Not rstDetalle.EOF
            If Me.rstDetalle!Modificar Then
                intIdMarcacion = Me.rstDetalle!idmarcacion
                ArrayTrabajador(intSeleccionados) = intIdMarcacion
                intSeleccionados = intSeleccionados + 1
            End If
            rstDetalle.MoveNext
        Loop
        ConfiguraArray = True
    End If

Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Set rstDetalle = Nothing
    ConfiguraArray = False

    Call ErrorMessage(FrmMantCompromisoMarcacion_ConfiguraArray, Err.Source & " FrmMantCompromisoMarcacion:ConfiguraArray", Err.Description)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mrstDetalle = Nothing
End Sub

Private Sub OptAmbas_Click()
    Me.LblHoraIngresoReal.Enabled = True
    Me.DtpHoraIngresoReal.Enabled = True
    Me.LblHoraSalidaReal.Enabled = True
    Me.DtpHoraSalidaReal.Enabled = True
End Sub

Private Sub OptNinguno_Click()
    Me.LblHoraIngresoReal.Enabled = False
    Me.DtpHoraIngresoReal.Enabled = False
    Me.LblHoraSalidaReal.Enabled = False
    Me.DtpHoraSalidaReal.Enabled = False
End Sub

Private Sub OptSoloIngreso_Click()
    Me.LblHoraIngresoReal.Enabled = True
    Me.DtpHoraIngresoReal.Enabled = True
    Me.LblHoraSalidaReal.Enabled = False
    Me.DtpHoraSalidaReal.Enabled = False
End Sub

Private Sub OptSoloSalida_Click()
    Me.LblHoraIngresoReal.Enabled = False
    Me.DtpHoraIngresoReal.Enabled = False
    Me.LblHoraSalidaReal.Enabled = True
    Me.DtpHoraSalidaReal.Enabled = True
End Sub

Private Sub TdbgAdministrar_Click()
If Me.TdbgAdministrar.Col = 6 Then
    If Me.rstDetalle.EOF And Me.rstDetalle.BOF Then
    If Me.rstDetalle!Modificar Then
        Me.rstDetalle!Modificar = False
    Else
        Me.rstDetalle!Modificar = True
    End If
    
    rstDetalle.Update
    End If
End If
End Sub

Public Function ValidateConsolidado() As Boolean
Dim objConsolidado As Object
On Error GoTo ErrHandler

    ValidateConsolidado = False
    Set objConsolidado = CreateObject("Personal.clsMovimientos")
    ValidateConsolidado = objConsolidado.GetExisteConsolidado(Me.DtpFechaInicio.Value)
    Set objConsolidado = Nothing

Exit Function
ErrHandler:
    Set objConsolidado = Nothing
    ValidateConsolidado = False
    Call ErrorMessage(FrmMantCompromisoMarcacionDelete_ValidateConsolidado, Err.Source & " FrmMantCompromisoMarcacionDelete:ValidateConsolidado", Err.Description)
End Function


