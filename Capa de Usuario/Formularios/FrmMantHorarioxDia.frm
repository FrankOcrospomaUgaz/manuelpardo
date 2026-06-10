VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmMantHorarioxDia 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12405
   Icon            =   "FrmMantHorarioxDia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   12405
   Begin VB.Frame FraDias 
      Caption         =   "Horario por día"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Left            =   4425
      TabIndex        =   39
      Top             =   1050
      Width           =   4005
      Begin VB.TextBox txtHoras 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   40
         Text            =   "0"
         Top             =   3960
         Width           =   1215
      End
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
         Left            =   3255
         Picture         =   "FrmMantHorarioxDia.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Quitar todos los registros"
         Top             =   1530
         Width           =   675
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
         Height          =   330
         Left            =   3255
         Picture         =   "FrmMantHorarioxDia.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Quitar el registro actual"
         Top             =   885
         Width           =   675
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
         Left            =   3255
         Picture         =   "FrmMantHorarioxDia.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Agregar registros"
         Top             =   285
         Width           =   675
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGDia 
         Height          =   3630
         Left            =   75
         TabIndex        =   19
         Top             =   255
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   6403
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Dia"
         Columns(0).DataField=   "Dia"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Ingreso"
         Columns(1).DataField=   "HoraInicio"
         Columns(1).NumberFormat=   "Medium Time"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Salida"
         Columns(2).DataField=   "HoraFin"
         Columns(2).NumberFormat=   "Medium Time"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1693"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=76696928"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1482"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1376"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=133636"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2752"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
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
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Horas:"
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
         Left            =   1320
         TabIndex        =   41
         Top             =   3960
         Width           =   540
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
      Height          =   4350
      Left            =   120
      TabIndex        =   29
      Top             =   1050
      Width           =   4275
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
         Left            =   135
         TabIndex        =   36
         Top             =   3510
         Width           =   4065
         Begin MSComCtl2.DTPicker DtpFechaFin 
            Height          =   330
            Left            =   2550
            TabIndex        =   15
            Top             =   270
            Width           =   1440
            _ExtentX        =   2540
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
            Format          =   17170433
            CurrentDate     =   38061
         End
         Begin MSComCtl2.DTPicker DtpFechaInicio 
            Height          =   330
            Left            =   645
            TabIndex        =   14
            Top             =   270
            Width           =   1440
            _ExtentX        =   2540
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
            Format          =   17170433
            CurrentDate     =   38061
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
            Left            =   105
            TabIndex        =   38
            Top             =   315
            Width           =   465
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
            Left            =   2235
            TabIndex        =   37
            Top             =   330
            Width           =   240
         End
      End
      Begin VB.Frame FraCriterio 
         Caption         =   "Seleccione Días"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   135
         TabIndex        =   35
         Top             =   1050
         Width           =   1545
         Begin VB.CheckBox chkDomingo 
            Caption         =   "Domingo"
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
            Left            =   165
            TabIndex        =   11
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CheckBox chkSabado 
            Caption         =   "Sábado"
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
            Left            =   165
            TabIndex        =   10
            Top             =   1740
            Width           =   975
         End
         Begin VB.CheckBox chkViernes 
            Caption         =   "Viernes"
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
            Left            =   165
            TabIndex        =   9
            Top             =   1455
            Width           =   975
         End
         Begin VB.CheckBox chkJueves 
            Caption         =   "Jueves"
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
            Left            =   165
            TabIndex        =   8
            Top             =   1170
            Width           =   975
         End
         Begin VB.CheckBox chkMiercoles 
            Caption         =   "Miércoles"
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
            Left            =   165
            TabIndex        =   7
            Top             =   870
            Width           =   1170
         End
         Begin VB.CheckBox chkMartes 
            Caption         =   "Martes"
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
            Left            =   165
            TabIndex        =   6
            Top             =   555
            Width           =   1125
         End
         Begin VB.CheckBox chkLunes 
            Caption         =   "Lunes"
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
            Left            =   165
            TabIndex        =   5
            Top             =   270
            Width           =   975
         End
      End
      Begin VB.Frame FraHora 
         Caption         =   "Rango de Horas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2385
         Left            =   1725
         TabIndex        =   32
         Top             =   1035
         Width           =   1575
         Begin MSComCtl2.DTPicker DtpHoraSalida 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   1635
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
            Format          =   17170435
            UpDown          =   -1  'True
            CurrentDate     =   38056
         End
         Begin MSComCtl2.DTPicker DtpHoraIngreso 
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   735
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
            Format          =   17170435
            UpDown          =   -1  'True
            CurrentDate     =   38056
         End
         Begin VB.Label LblHoraSalida 
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
            Left            =   135
            TabIndex        =   34
            Top             =   1305
            Width           =   510
         End
         Begin VB.Label LblHoraIngreso 
            AutoSize        =   -1  'True
            Caption         =   "Ingreso"
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
            TabIndex        =   33
            Top             =   390
            Width           =   660
         End
      End
      Begin VB.Frame FraTipoHora 
         Caption         =   "Tipo Hora"
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
         TabIndex        =   30
         Top             =   240
         Width           =   4065
         Begin MSDataListLib.DataCombo DtcTipoHora 
            Height          =   315
            Left            =   1455
            TabIndex        =   4
            Top             =   270
            Width           =   2550
            _ExtentX        =   4498
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
         Begin VB.Label LblDescripcion 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
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
            TabIndex        =   31
            Top             =   330
            Width           =   975
         End
      End
      Begin VB.CommandButton CmdAgregarDia 
         Height          =   330
         Left            =   3390
         Picture         =   "FrmMantHorarioxDia.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Agregar registros"
         Top             =   1125
         Width           =   705
      End
      Begin VB.CommandButton CmdQuitarDia 
         Height          =   330
         Left            =   3390
         Picture         =   "FrmMantHorarioxDia.frx":19D2
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Quitar el registro actual"
         Top             =   1725
         Width           =   705
      End
      Begin VB.CommandButton CmdQuitarTodosDias 
         Height          =   330
         Left            =   3390
         Picture         =   "FrmMantHorarioxDia.frx":1E14
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Quitar todos los registros"
         Top             =   2325
         Width           =   705
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
      Left            =   120
      TabIndex        =   28
      Top             =   5385
      Width           =   8325
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   3435
         Picture         =   "FrmMantHorarioxDia.frx":2256
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Confirmar operación"
         Top             =   195
         Width           =   810
      End
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   4335
         Picture         =   "FrmMantHorarioxDia.frx":3298
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Cancelar la operación"
         Top             =   195
         Width           =   810
      End
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
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   8295
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
         Left            =   7650
         TabIndex        =   3
         ToolTipText     =   "Búsqueda de Trabajador"
         Top             =   270
         Width           =   465
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
         Left            =   1785
         MaxLength       =   60
         TabIndex        =   2
         Tag             =   "-1"
         Top             =   615
         Width           =   6375
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
         Left            =   1785
         TabIndex        =   1
         Tag             =   "-1"
         Text            =   "0"
         Top             =   210
         Width           =   975
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
         TabIndex        =   27
         Top             =   660
         Width           =   1530
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
         TabIndex        =   26
         Top             =   270
         Width           =   570
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGHorario 
      Height          =   6105
      Left            =   8460
      TabIndex        =   23
      Top             =   90
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   10769
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1826"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=56,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
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
End
Attribute VB_Name = "FrmMantHorarioxDia"
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

Private Array1() As String
Public strCadenaHorario As String
'********************Private Array2() As String
Private ArrayTrabajador() As String
Private mrstTipoHora As Object
Private mrstDetalle As Object
Private mrstDias As Object
Private intIdDia As Integer

Private Const mstrformname As String = "Mantenimiento de Horario x Trabajador x Días"

Public Property Set rstDetalle(ByVal vNewValue As Variant)
    Set mrstDetalle = vNewValue
End Property

Public Property Set rstDias(ByVal vNewValue As Variant)
    Set mrstDias = vNewValue
End Property

Public Property Get rstDetalle() As Variant
    Set rstDetalle = mrstDetalle
End Property

Public Property Get rstDias() As Variant
    Set rstDias = mrstDias
End Property

Public Property Get rstTipoHora() As Variant
    Set rstTipoHora = mrstTipoHora
End Property
Public Property Set rstTipoHora(ByVal vNewValue As Variant)
    Set mrstTipoHora = vNewValue
    If Not rstTipoHora.EOF Then
        Set Me.DtcTipoHora.RowSource = rstTipoHora
        Me.DtcTipoHora.ListField = "Descripcion"
        Me.DtcTipoHora.BoundColumn = "IdTipoHora"
        Me.DtcTipoHora.Enabled = True
    Else
        Me.DtcTipoHora.Enabled = False
        Me.DtcTipoHora.Text = ""
    End If
End Property

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
            Me.DtcTipoHora.BoundText = 0
            Me.DtpFechaInicio.Value = Date
            Me.DtpFechaFin.Value = Date
            Me.DtpHoraIngreso.Value = Time
            Me.DtpHoraSalida.Value = Time
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
        If ActiveForm("FrmConsHorarioxDia") Then
            FrmConsHorarioxDia.LoadData
        End If
    End If

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantHorarioxTrabajador_ShowForm, Err.Source & " FrmMantHorarioxTrabajador:ShowForm", Err.Description)
End Sub

Private Sub CmdAgregar_Click()
Dim datFechaInicio As Date
Dim datFechaFin As Date
Dim objDias As Object
    If Me.TxtIdTrabajador.Text <> 0 Or Me.DtcTipoHora.BoundText <> "" Or Me.DtpFechaInicio > Me.DtpFechaFin Then
        rstDetalle.Filter = "IdTrabajador = " & Val(Me.TxtIdTrabajador.Text)
        Set objDias = mrstDias.Clone
        datFechaInicio = Me.DtpFechaInicio.Value
        datFechaFin = Me.DtpFechaFin.Value
        If objDias.RecordCount > 0 Then
            If ValidateFieldsAgregar Then
                Do While datFechaInicio <= datFechaFin
                strCadenaHorario = ""
                objDias.MoveFirst
                    Do While Not objDias.EOF
                        strCadenaHorario = strCadenaHorario + IIf(strCadenaHorario = "", "", "?") + objDias!Dia + "@" + CStr(objDias!horaInicio) + "@" + CStr(objDias!horaFin)
                        If objDias!Dia = "Lunes" And WeekDay(datFechaInicio) = 2 Then
                            rstDetalle.AddNew
                            rstDetalle!IdTrabajador = Val(Me.TxtIdTrabajador.Text)
                            rstDetalle!FECHA = datFechaInicio
                            rstDetalle!HoraIngreso = Format(objDias!horaInicio, "hh:mm:00")
                            rstDetalle!HoraSalida = Format(objDias!horaFin, "hh:mm:00")
                            rstDetalle.Update
                        End If
                        Dim a As Variant
                        a = WeekDay(datFechaInicio)
                        If objDias!Dia = "Martes" And WeekDay(datFechaInicio) = 3 Then
                            rstDetalle.AddNew
                            rstDetalle!IdTrabajador = Val(Me.TxtIdTrabajador.Text)
                            rstDetalle!FECHA = datFechaInicio
                            rstDetalle!HoraIngreso = Format(objDias!horaInicio, "hh:mm:00")
                            rstDetalle!HoraSalida = Format(objDias!horaFin, "hh:mm:00")
                            rstDetalle.Update
                        End If
                        If objDias!Dia = "Miércoles" And WeekDay(datFechaInicio) = 4 Then
                            rstDetalle.AddNew
                            rstDetalle!IdTrabajador = Val(Me.TxtIdTrabajador.Text)
                            rstDetalle!FECHA = datFechaInicio
                            rstDetalle!HoraIngreso = Format(objDias!horaInicio, "hh:mm:00")
                            rstDetalle!HoraSalida = Format(objDias!horaFin, "hh:mm:00")
                            rstDetalle.Update
                        End If
                        If objDias!Dia = "Jueves" And WeekDay(datFechaInicio) = 5 Then
                            rstDetalle.AddNew
                            rstDetalle!IdTrabajador = Val(Me.TxtIdTrabajador.Text)
                            rstDetalle!FECHA = datFechaInicio
                            rstDetalle!HoraIngreso = Format(objDias!horaInicio, "hh:mm:00")
                            rstDetalle!HoraSalida = Format(objDias!horaFin, "hh:mm:00")
                            rstDetalle.Update
                        End If
                        If objDias!Dia = "Viernes" And WeekDay(datFechaInicio) = 6 Then
                            rstDetalle.AddNew
                            rstDetalle!IdTrabajador = Val(Me.TxtIdTrabajador.Text)
                            rstDetalle!FECHA = datFechaInicio
                            rstDetalle!HoraIngreso = Format(objDias!horaInicio, "hh:mm:00")
                            rstDetalle!HoraSalida = Format(objDias!horaFin, "hh:mm:00")
                            rstDetalle.Update
                        End If
                        If objDias!Dia = "Sábado" And WeekDay(datFechaInicio) = 7 Then
                            rstDetalle.AddNew
                            rstDetalle!IdTrabajador = Val(Me.TxtIdTrabajador.Text)
                            rstDetalle!FECHA = datFechaInicio
                            rstDetalle!HoraIngreso = Format(objDias!horaInicio, "hh:mm:00")
                            rstDetalle!HoraSalida = Format(objDias!horaFin, "hh:mm:00")
                            rstDetalle.Update
                        End If
                        If objDias!Dia = "Domingo" And WeekDay(datFechaInicio) = 1 Then
                            rstDetalle.AddNew
                            rstDetalle!IdTrabajador = Val(Me.TxtIdTrabajador.Text)
                            rstDetalle!FECHA = datFechaInicio
                            rstDetalle!HoraIngreso = Format(objDias!horaInicio, "hh:mm:00")
                            rstDetalle!HoraSalida = Format(objDias!horaFin, "hh:mm:00")
                            rstDetalle.Update
                        End If
                        objDias.MoveNext
                    Loop
                    datFechaInicio = datFechaInicio + 1
                Loop
            End If
        End If
        
          
        Set Me.TDBGHorario.DataSource = mrstDetalle
        Me.TDBGHorario.ReBind
        
    Else
        MsgBox MSG_SELECTION_EMPTY, vbExclamation, MSG_ACTION_ADDNEW & " - " & mstrformname
    End If
End Sub

Private Sub CmdAgregarDia_Click()
Dim objDias As Object
Dim intTotalMinutos As Double
If Me.TxtIdTrabajador.Text <> 0 Or Me.DtcTipoHora.BoundText <> "" Or Me.DtpFechaInicio > Me.DtpFechaFin Then
If Not ValidateTime() Then Exit Sub
        rstDetalle.Filter = "IdTrabajador = " & Val(Me.TxtIdTrabajador.Text)
        If Me.chkLunes.Value = 1 Then
            rstDias.AddNew
            intIdDia = intIdDia + 1
            rstDias!IdDia = intIdDia
            rstDias!Dia = "Lunes"
            rstDias!horaInicio = Me.DtpHoraIngreso
            rstDias!horaFin = Me.DtpHoraSalida
            rstDias.Update
        End If
        If Me.chkMartes.Value = 1 Then
            rstDias.AddNew
            intIdDia = intIdDia + 1
            rstDias!IdDia = intIdDia
            rstDias!Dia = "Martes"
            rstDias!horaInicio = Me.DtpHoraIngreso
            rstDias!horaFin = Me.DtpHoraSalida
            rstDias.Update
        End If
        If Me.chkMiercoles.Value = 1 Then
            rstDias.AddNew
            intIdDia = intIdDia + 1
            rstDias!IdDia = intIdDia
            rstDias!Dia = "Miércoles"
            rstDias!horaInicio = Me.DtpHoraIngreso
            rstDias!horaFin = Me.DtpHoraSalida
            rstDias.Update
        End If
        If Me.chkJueves.Value = 1 Then
            rstDias.AddNew
            intIdDia = intIdDia + 1
            rstDias!IdDia = intIdDia
            rstDias!Dia = "Jueves"
            rstDias!horaInicio = Me.DtpHoraIngreso
            rstDias!horaFin = Me.DtpHoraSalida
            rstDias.Update
        End If
        If Me.chkViernes.Value = 1 Then
            rstDias.AddNew
            intIdDia = intIdDia + 1
            rstDias!IdDia = intIdDia
            rstDias!Dia = "Viernes"
            rstDias!horaInicio = Me.DtpHoraIngreso
            rstDias!horaFin = Me.DtpHoraSalida
            rstDias.Update
        End If
        If Me.chkSabado.Value = 1 Then
            rstDias.AddNew
            intIdDia = intIdDia + 1
            rstDias!IdDia = intIdDia
            rstDias!Dia = "Sábado"
            rstDias!horaInicio = Me.DtpHoraIngreso
            rstDias!horaFin = Me.DtpHoraSalida
            rstDias.Update
        End If
        If Me.chkDomingo.Value = 1 Then
            rstDias.AddNew
            intIdDia = intIdDia + 1
            rstDias!IdDia = intIdDia
            rstDias!Dia = "Domingo"
            rstDias!horaInicio = Me.DtpHoraIngreso
            rstDias!horaFin = Me.DtpHoraSalida
            rstDias.Update
        End If
        
        Set objDias = mrstDias.Clone
        objDias.MoveFirst
        intTotalMinutos = 0
        Do While Not objDias.EOF
            intTotalMinutos = intTotalMinutos + DateDiff("n", objDias!horaInicio, objDias!horaFin)
            objDias.MoveNext
        Loop
        Me.txtHoras.Text = intTotalMinutos / 60
        
        Set Me.TDBGDia.DataSource = mrstDias
        Me.TDBGDia.ReBind
Else
    MsgBox MSG_SELECTION_EMPTY, vbExclamation, MSG_ACTION_ADDNEW & " - " & mstrformname
End If
End Sub

Private Sub CmdBuscarTrabajador_Click()
Dim miArray() As String
Dim miCampos() As String
Dim intX As Integer
Dim intTotalMinutos As Double
Dim objDias As Object
    FrmBusqTrabajador.strModo = Find
    FrmBusqTrabajador.ShowForm

    If Trim(FrmBusqTrabajador.intIdTrabajador) <> 0 Then
        Me.TxtIdTrabajador.Text = FrmBusqTrabajador.intIdTrabajador
        Me.strCadenaHorario = FrmBusqTrabajador.strCadenaHorario
        miArray = Split(strCadenaHorario, "?")
        For intX = 0 To UBound(miArray)
            miCampos = Split(miArray(intX), "@")
            rstDias.AddNew
            rstDias!IdDia = intX
            rstDias!Dia = miCampos(0)
            rstDias!horaInicio = miCampos(1)
            rstDias!horaFin = miCampos(2)
            rstDias.Update
        Next
        
        Set objDias = mrstDias.Clone
        If objDias.RecordCount Then
        objDias.MoveFirst
        intTotalMinutos = 0
        Do While Not objDias.EOF
            intTotalMinutos = intTotalMinutos + DateDiff("n", objDias!horaInicio, objDias!horaFin)
            objDias.MoveNext
        Loop
        Me.txtHoras.Text = intTotalMinutos / 60
        Else
        Me.txtHoras.Text = "0"
        End If
        Set Me.TDBGDia.DataSource = mrstDias
        Me.TDBGDia.ReBind
        
        Me.TxtApellido.Text = Trim$(FrmBusqTrabajador.strApellidoNombre)
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

Private Sub CmdQuitarDia_Click()
Dim datFecha As Date
Dim strDia As String
Dim objDias As Object
Dim intTotalMinutos As Double
    If Me.TDBGDia.SelBookmarks.Count = 0 Then
        MsgBox MSG_SELECTION_EMPTY, vbExclamation, MSG_ACTION_DELETE & " - " & mstrformname
    Else
        strDia = CStr(rstDias!IdDia)
        rstDias.Filter = "IdDia =" & strDia
        rstDias.Delete
        rstDias.Filter = ""
        
        Set objDias = mrstDias.Clone
        objDias.MoveFirst
        intTotalMinutos = 0
        Do While Not objDias.EOF
            intTotalMinutos = intTotalMinutos + DateDiff("n", objDias!horaInicio, objDias!horaFin)
            objDias.MoveNext
        Loop
        Me.txtHoras.Text = intTotalMinutos / 60
        
        Set Me.TDBGDia.DataSource = mrstDias
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

Private Sub CmdQuitarTodosDias_Click()
Dim objPersonal As Object

    Set objPersonal = CreateObject("Personal.clsGeneral")
    Set rstDias = objPersonal.GetCampos("Entero AS IdDia, Texto AS Dia, HoraLarga as HoraInicio, HoraLarga as HoraFin")
    Set objPersonal = Nothing
    
    Me.txtHoras.Text = "0"
    
    Set Me.TDBGDia.DataSource = mrstDias
    Me.TDBGHorario.ReBind
    intIdDia = 0
Exit Sub
ErrHandler:
    Set objPersonal = Nothing

    Call ErrorMessage(FrmMantHorarioxTrabajador_LoadData, Err.Source & " FrmMantHorarioxTrabajador:LoadData", Err.Description)
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.strMensajeError = ""
End Sub

Private Sub LoadData()
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set rstTipoHora = objPersonal.GetTipoHora()
    Set objPersonal = Nothing

    Set objPersonal = CreateObject("Personal.clsGeneral")
    Set rstDias = objPersonal.GetCampos("Entero AS IdDia, Texto as Dia, HoraLarga as HoraInicio, HoraLarga as HoraFin")
    Set rstDetalle = objPersonal.GetCampos("Entero As IdTrabajador, FechaMediana as Fecha, HoraLarga As HoraIngreso, HoraLarga as HoraSalida")
       
    Set objPersonal = Nothing
    
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Call ErrorMessage(FrmMantHorarioxTrabajador_LoadData, Err.Source & " FrmMantHorarioxTrabajador:LoadData", Err.Description)
End Sub

Private Function AddNew() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    
    Set objPersonal = CreateObject("Personal.clsHorario")
    AddNew = objPersonal.AddNewHorarioxTrabajador(Array1(), ArrayTrabajador())
    Set objPersonal = Nothing
    If AddNew Then
        Bitacora "Horario", KEY_ACTION_ADDNEW
    End If
    
Exit Function
ErrHandler:
    Set objPersonal = Nothing
    AddNew = False

Call ErrorMessage(FrmMantHorarioxTrabajador_AddNew, Err.Source & " FrmMantHorarioxTrabajador:AddNew", Err.Description)
End Function
Private Function Update() As Boolean

End Function

Public Function ValidateFieldsAgregar() As Boolean
    ValidateFieldsAgregar = False
    
    If Me.DtpHoraSalida.Value - Me.DtpHoraIngreso.Value < 0 Then
        MsgBox MSG_INVALIDATE_DATE, vbInformation, mstrformname
        DtpHoraIngreso.SetFocus
        Exit Function
    End If
        
    If Me.DtpFechaFin.Value - Me.DtpFechaInicio.Value < 0 Then
        MsgBox MSG_INVALIDATE_DATE, vbInformation, mstrformname
        DtpHoraIngreso.SetFocus
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
    If Me.DtcTipoHora.BoundText = "" Then
        MsgBox MSG_EMPTY_TIPOHORA, vbInformation, mstrformname
        Me.DtcTipoHora.SetFocus
        Exit Function
    End If
    Me.intIdTrabajador = Val(Me.TxtIdTrabajador.Text)
    Me.intIdTipoHora = Me.DtcTipoHora.BoundText
    If ConfiguraArray Then
        ValidateFields = True
    End If
    
End Function

Private Function ConfiguraArray() As Boolean
Dim intContador As Integer
Dim dblTotalHoras As Double
Dim datFechaTmp As Date
On Error GoTo ErrHandler

    intContador = 0
    ReDim Array1(0 To rstDetalle.RecordCount * 2)
   '******************* ReDim Array2(0 To rstDetalle.RecordCount)
    ReDim ArrayTrabajador(0 To 2)
    ArrayTrabajador(0) = "up_ModificarTrabajadorxCompromiso '" & intIdTrabajador & "','" & gstrCompromisoNoGenerado & "','" & strCadenaHorario & "'"
    rstDetalle.MoveFirst
    Do While Not rstDetalle.EOF
        dblTotalHoras = DateDiff("n", rstDetalle!HoraIngreso, rstDetalle!HoraSalida)
        
        If datFechaTmp <> rstDetalle!FECHA Then
            datFechaTmp = rstDetalle!FECHA
            Array1(intContador) = "up_EliminarHorarioxTrabajadorGenerado '" & intIdTrabajador & "', '" & rstDetalle!FECHA & "'"
            intContador = intContador + 1
        End If
        
        Array1(intContador) = "up_AgregarHorarioxTrabajador '" & Me.intIdTipoHora & "','" & intIdTrabajador & "', '" & rstDetalle!FECHA & "','" & FormatDateTime(rstDetalle!HoraIngreso, vbShortTime) & "','" & FormatDateTime(rstDetalle!HoraSalida, vbShortTime) & "','" & dblTotalHoras & "'"
        'Array1(intContador) = "INSERT INTO HorarioxTrabajador (IdHorario, IdTipoHora, IdTrabajador, Fecha, HoraIngreso, HoraSalida, TotalHoras) VALUES ("
        'Array2(intContador) = ", " & Me.intIdTipoHora & ", " & Me.intIdTrabajador & ", '" & rstDetalle!Fecha & "','" & rstDetalle!HoraIngreso & _
                    "','" & rstDetalle!HoraSalida & "', " & dblTotalHoras & ")"
        rstDetalle.MoveNext
        intContador = intContador + 1
    Loop
    ConfiguraArray = True

Exit Function
ErrHandler:
    ConfiguraArray = False

Call ErrorMessage(FrmMantHorarioxTrabajador_ConfiguraArray, Err.Source & " FrmMantHorarioxTrabajador:ConfiguraArray", Err.Description)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mrstTipoHora = Nothing
    Set mrstDetalle = Nothing
    Set mrstDias = Nothing
End Sub

Private Function ValidateTime() As Boolean
    ValidateTime = False
    If Me.DtpHoraIngreso.Value >= Me.DtpHoraSalida Then
        MsgBox MSG_INVALIDATE_DATE, vbInformation, mstrformname
        Exit Function
    End If
    ValidateTime = True
End Function
