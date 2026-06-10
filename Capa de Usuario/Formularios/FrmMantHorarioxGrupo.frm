VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmMantHorarioxGrupo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   Icon            =   "FrmMantHorarioxGrupo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8970
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
      Left            =   45
      TabIndex        =   19
      Top             =   795
      Width           =   4920
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
         TabIndex        =   26
         Top             =   2970
         Width           =   4680
         Begin MSComCtl2.DTPicker DtpFechaFin 
            Height          =   330
            Left            =   3255
            TabIndex        =   8
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
            Format          =   21364737
            CurrentDate     =   38061
         End
         Begin MSComCtl2.DTPicker DtpFechaInicio 
            Height          =   330
            Left            =   840
            TabIndex        =   7
            Top             =   270
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
            Format          =   21364737
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
            Left            =   150
            TabIndex        =   28
            Top             =   300
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
            Left            =   2640
            TabIndex        =   27
            Top             =   345
            Width           =   240
         End
      End
      Begin VB.Frame FraCriterio 
         Caption         =   "Criterios de creación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   120
         TabIndex        =   25
         Top             =   1155
         Width           =   3735
         Begin VB.OptionButton OptTodos 
            Caption         =   "Todos los días"
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
            Left            =   135
            TabIndex        =   3
            Top             =   255
            Width           =   3105
         End
         Begin VB.OptionButton OptDomingo 
            Caption         =   "Excluir días Domingo"
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
            Left            =   135
            TabIndex        =   4
            Top             =   600
            Width           =   3105
         End
         Begin VB.OptionButton OptSabadoDomingo 
            Caption         =   "Excluir días Sábado y Domingo"
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
            Left            =   135
            TabIndex        =   5
            Top             =   960
            Width           =   3105
         End
         Begin VB.OptionButton OptContinuo 
            Caption         =   "Continuo"
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
            TabIndex        =   6
            Top             =   1320
            Width           =   2730
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
         Height          =   645
         Left            =   120
         TabIndex        =   22
         Top             =   3795
         Width           =   4650
         Begin MSComCtl2.DTPicker DtpHoraSalida 
            Height          =   285
            Left            =   3225
            TabIndex        =   10
            Top             =   255
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
            Format          =   21364739
            UpDown          =   -1  'True
            CurrentDate     =   38056
         End
         Begin MSComCtl2.DTPicker DtpHoraIngreso 
            Height          =   285
            Left            =   870
            TabIndex        =   9
            Top             =   255
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
            Format          =   21364739
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
            Left            =   2595
            TabIndex        =   24
            Top             =   285
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
            Left            =   90
            TabIndex        =   23
            Top             =   285
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
         TabIndex        =   20
         Top             =   285
         Width           =   4710
         Begin MSDataListLib.DataCombo DtcTipoHora 
            Height          =   315
            Left            =   1290
            TabIndex        =   2
            Top             =   255
            Width           =   3360
            _ExtentX        =   5927
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
            Left            =   150
            TabIndex        =   21
            Top             =   345
            Width           =   975
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
         Height          =   330
         Left            =   4035
         Picture         =   "FrmMantHorarioxGrupo.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Agregar registros"
         Top             =   1260
         Width           =   705
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
         Left            =   4035
         Picture         =   "FrmMantHorarioxGrupo.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Quitar el registro actual"
         Top             =   1830
         Width           =   705
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
         Left            =   4005
         Picture         =   "FrmMantHorarioxGrupo.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Quitar todos los registros"
         Top             =   2460
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
      Left            =   60
      TabIndex        =   15
      Top             =   5430
      Width           =   4920
      Begin VB.CommandButton CmdAceptar 
         Height          =   510
         Left            =   1650
         Picture         =   "FrmMantHorarioxGrupo.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Confirmar operación"
         Top             =   180
         Width           =   810
      End
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   2580
         Picture         =   "FrmMantHorarioxGrupo.frx":25D2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Cancelar la operación"
         Top             =   180
         Width           =   810
      End
   End
   Begin VB.Frame FraTrabajador 
      Caption         =   "Grupo de Trabajador"
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
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   4920
      Begin MSDataListLib.DataCombo DtcGrupo 
         Height          =   315
         Left            =   1395
         TabIndex        =   1
         Top             =   225
         Width           =   3360
         _ExtentX        =   5927
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
      Begin VB.Label LblGrupo 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
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
         Left            =   270
         TabIndex        =   14
         Top             =   300
         Width           =   510
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGHorario 
      Height          =   6090
      Left            =   5040
      TabIndex        =   18
      Top             =   135
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   10742
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
Attribute VB_Name = "FrmMantHorarioxGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
'Public intIdHorario As Long
Public intIdGrupo As Integer
Public intIdTipoHora As Integer
Public datFecha As Date
'Public datHoraIngreso As Date
'Public datHoraSalida As Date
Public blnStatus As Boolean
Public strMensajeError As String

Private Array1() As String

Private ArrayTrabajador() As String

Private mrstTipoHora As Object
Private mrstGrupo As Object
Private mrstDetalle As Object

Private Const mstrformname As String = "Mantenimiento de Horario x Grupo"

Public Property Set rstDetalle(ByVal vNewValue As Variant)
    Set mrstDetalle = vNewValue
End Property
Public Property Get rstDetalle() As Variant
    Set rstDetalle = mrstDetalle
End Property

Public Property Get rstTipoHora() As Variant
    Set rstTipoHora = mrstTipoHora
End Property
Public Property Set rstTipoHora(ByVal vNewValue As Variant)
    Set mrstTipoHora = vNewValue
    If Not rstTipoHora.EOF Then
        Set Me.dtcTipoHora.RowSource = rstTipoHora
        Me.dtcTipoHora.ListField = "Descripcion"
        Me.dtcTipoHora.BoundColumn = "IdTipoHora"
        Me.dtcTipoHora.Enabled = True
    Else
        Me.dtcTipoHora.Enabled = False
        Me.dtcTipoHora.Text = ""
    End If
End Property

Public Property Get rstGrupo() As Variant
    Set rstGrupo = mrstGrupo
End Property
Public Property Set rstGrupo(ByVal vNewValue As Variant)
    Set mrstGrupo = vNewValue
    If Not rstGrupo.EOF Then
        Set Me.dtcGrupo.RowSource = rstGrupo
        Me.dtcGrupo.ListField = "Descripcion"
        Me.dtcGrupo.BoundColumn = "IdGrupo"
        Me.dtcGrupo.Enabled = True
    Else
        Me.dtcGrupo.Enabled = False
        Me.dtcGrupo.Text = ""
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
            Me.dtcGrupo.BoundText = 0
            Me.dtcTipoHora.BoundText = 0
            Me.DtpFechaInicio.Value = Date
            Me.DtpFechaFin.Value = Date
            Me.DtpHoraIngreso.Value = Time
            Me.DtpHoraSalida.Value = Time
            Me.OptSabadoDomingo.Value = True
    End Select
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantHorarioxGrupo_ShowForm, Err.Source & " FrmMantHorarioxGrupo:ShowForm", Err.Description)
End Sub

Private Sub CmdAceptar_Click()
Dim blnResult As Boolean
Screen.MousePointer = vbHourglass
On Error GoTo ErrHandler
    If Not ValidateFields Then Exit Sub
    Select Case intAction
        Case enumActions.Update
            blnResult = Update()
        Case enumActions.AddNew
            blnResult = AddNew()
    End Select
    
    If Not blnResult Then
        Screen.MousePointer = vbDefault
        MsgBox MSG_OPERATION_FAIL & vbCrLf & Me.strMensajeError, vbCritical, mstrformname
        Me.blnStatus = False
    Else
        Unload Me
        If ActiveForm("FrmConsHorarioxGrupo") Then
            FrmConsHorarioxGrupo.LoadData
        End If
    End If
    Screen.MousePointer = vbDefault
Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    Call ErrorMessage(FrmMantHorarioxGrupo_ShowForm, Err.Source & " FrmMantHorarioxGrupo:ShowForm", Err.Description)
End Sub

Private Sub CmdAgregar_Click()
Dim datFechaInicio As Date
Dim datFechaFin As Date

    If Me.dtcGrupo.BoundText <> "" Or Me.dtcTipoHora.BoundText <> "" Or Me.DtpFechaInicio > Me.DtpFechaFin Then
        datFechaInicio = Me.DtpFechaInicio.Value
        datFechaFin = Me.DtpFechaFin.Value
        
        If Me.OptTodos.Value Then
            If ValidateFieldsAgregar Then
                Do While datFechaInicio <= datFechaFin
                    rstDetalle.AddNew
                    rstDetalle!Fecha = datFechaInicio
                    rstDetalle!HoraIngreso = FormatDateTime(Me.DtpHoraIngreso.Value, vbShortTime)
                    rstDetalle!HoraSalida = FormatDateTime(Me.DtpHoraSalida.Value, vbShortTime)
                    rstDetalle.Update
                    datFechaInicio = datFechaInicio + 1
                Loop
            End If
        End If
        
        If Me.OptDomingo.Value Then
            If ValidateFieldsAgregar Then
                Do While datFechaInicio <= datFechaFin
                    If WeekDay(datFechaInicio) <> 1 Then
                        rstDetalle.AddNew
                        rstDetalle!Fecha = datFechaInicio
                        rstDetalle!HoraIngreso = FormatDateTime(Me.DtpHoraIngreso.Value, vbShortTime)
                        rstDetalle!HoraSalida = FormatDateTime(Me.DtpHoraSalida.Value, vbShortTime)
                        rstDetalle.Update
                    End If
                    datFechaInicio = datFechaInicio + 1
                Loop
            End If
        End If
        
        If Me.OptSabadoDomingo.Value Then
            If ValidateFieldsAgregar Then
                Do While datFechaInicio <= datFechaFin
                    If WeekDay(datFechaInicio) <> 1 And WeekDay(datFechaInicio) <> 7 Then
                        rstDetalle.AddNew
                        rstDetalle!Fecha = datFechaInicio
                        rstDetalle!HoraIngreso = Format(Me.DtpHoraIngreso.Value, "hh:mm:00")
                        rstDetalle!HoraSalida = Format(Me.DtpHoraSalida.Value, "hh:mm:00")
                        rstDetalle.Update
                    End If
                    datFechaInicio = datFechaInicio + 1
                Loop
            End If
        End If
        
        If Me.OptContinuo.Value Then
            If ValidateFieldsAgregar Then
                'HorarioContinuoGrupo rstDetalle, Me.TxtIdTrabajador.Text, datFechaInicio, _
                                 datFechaFin, Me.DtpHoraIngreso.Value, Me.DtpHoraSalida.Value
            End If
        End If
                
        Set Me.TDBGHorario.DataSource = mrstDetalle
        Me.TDBGHorario.ReBind
        
    Else
        MsgBox MSG_SELECTION_EMPTY, vbExclamation, MSG_ACTION_ADDNEW & " - " & mstrformname
    End If
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdQuitar_Click()
Dim datFecha As Date

    If Me.TDBGHorario.SelBookmarks.Count = 0 Then
        MsgBox MSG_SELECTION_EMPTY, vbExclamation, MSG_ACTION_DELETE & " - " & mstrformname
    Else
        datFecha = rstDetalle!Fecha
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

    Call ErrorMessage(FrmMantHorarioxGrupo_LoadData, Err.Source & " FrmMantHorarioxGrupo:LoadData", Err.Description)
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

    Set objPersonal = CreateObject("Personal.clsPersonal")
    Set rstGrupo = objPersonal.GetGrupo()
    Set objPersonal = Nothing
    
    Set objPersonal = CreateObject("Personal.clsGeneral")
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

Call ErrorMessage(FrmMantHorarioxGrupo_AddNew, Err.Source & " FrmMantHorarioxGrupo:AddNew", Err.Description)
End Function
Private Function Update() As Boolean

End Function

Public Function ValidateFieldsAgregar() As Boolean
    ValidateFieldsAgregar = False
    
    If Not (Me.OptContinuo.Value) And Me.DtpHoraSalida.Value - Me.DtpHoraIngreso.Value < 0 Then
        MsgBox MSG_INVALIDATE_DATE, vbInformation, mstrformname
        DtpHoraIngreso.SetFocus
        Exit Function
    End If
        
    If Me.OptContinuo.Value And Me.DtpFechaFin.Value - Me.DtpFechaInicio.Value < 0 Then
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
    
    Me.intIdGrupo = Me.dtcGrupo.BoundText
    Me.intIdTipoHora = Me.dtcTipoHora.BoundText
    If ConfiguraArray Then
        ValidateFields = True
    End If
    
End Function

Private Function ConfiguraArray() As Boolean
Dim intContador As Long
Dim intContadorTrab As Integer
Dim dblTotalHoras As Double
Dim intIdTrabajador As Integer
Dim objPersonal As Object
Dim rstTrabajador As Object

On Error GoTo ErrHandler

'selecciona todos los trabajadores q pertenecen al grupo
Set objPersonal = CreateObject("Personal.clsPersonal")
Set rstTrabajador = objPersonal.GetTrabajador(, , , , , , , , , , , , , , Me.intIdGrupo)
Set objPersonal = Nothing

intContador = 0
intContadorTrab = 0
ReDim Array1(0 To rstDetalle.RecordCount * rstTrabajador.RecordCount)
'**************************ReDim Array2(0 To rstDetalle.RecordCount * rstTrabajador.RecordCount)
ReDim ArrayTrabajador(0 To rstTrabajador.RecordCount)

If Not rstTrabajador.BOF Or Not rstTrabajador.EOF Then
    rstTrabajador.MoveFirst
    Do While Not rstTrabajador.EOF
        intIdTrabajador = rstTrabajador!IdTrabajador
        ArrayTrabajador(intContadorTrab) = "up_ModificarTrabajadorxCompromiso '" & intIdTrabajador & "','" & gstrCompromisoNoGenerado & "'"
        intContadorTrab = intContadorTrab + 1
        rstDetalle.MoveFirst
        Do While Not rstDetalle.EOF
            dblTotalHoras = DateDiff("n", rstDetalle!HoraIngreso, rstDetalle!HoraSalida)
            Array1(intContador) = "up_AgregarHorarioxTrabajador '" & Me.intIdTipoHora & "','" & intIdTrabajador & "', '" & rstDetalle!Fecha & "','" & FormatDateTime(rstDetalle!HoraIngreso, vbShortTime) & "','" & FormatDateTime(rstDetalle!HoraSalida, vbShortTime) & "','" & dblTotalHoras & "'"
            'Array1(intContador) = "INSERT INTO HorarioxTrabajador (IdHorario, IdTipoHora, IdTrabajador, Fecha, HoraIngreso, HoraSalida, TotalHoras) VALUES ("
            'Array2(intContador) = ", " & Me.intIdTipoHora & ", " & intIdTrabajador & ", '" & rstDetalle!Fecha & "','" & rstDetalle!HoraIngreso & _
            '           "','" & rstDetalle!HoraSalida & "', " & dblTotalHoras & ")"
            rstDetalle.MoveNext
            intContador = intContador + 1
        Loop
        rstTrabajador.MoveNext
    Loop
End If
ConfiguraArray = True

Exit Function
ErrHandler:
    ConfiguraArray = False
    Set objPersonal = Nothing
    Set rstTrabajador = Nothing

Call ErrorMessage(FrmMantHorarioxGrupo_ConfiguraArray, Err.Source & " FrmMantHorarioxGrupo:ConfiguraArray", Err.Description)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mrstTipoHora = Nothing
    Set mrstDetalle = Nothing
    Set mrstGrupo = Nothing
End Sub


