VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmMantCompromisoMarcacionDelete 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   Icon            =   "FrmMantCompromisoMarcacionDelete.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8085
   Begin VB.Frame FraFecha 
      Caption         =   "Desde la Fecha"
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
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   7860
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
         Left            =   6585
         Picture         =   "FrmMantCompromisoMarcacionDelete.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Consultar inasistencias por fecha"
         Top             =   210
         Width           =   630
      End
      Begin MSComCtl2.DTPicker DtpFechaInicio 
         Height          =   285
         Left            =   3600
         TabIndex        =   7
         Top             =   315
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
         Format          =   63766529
         CurrentDate     =   38056
      End
      Begin VB.Label LblFechaInicio 
         AutoSize        =   -1  'True
         Caption         =   "Filtrar Desde :"
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
         Left            =   1800
         TabIndex        =   8
         Top             =   360
         Width           =   1170
      End
   End
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
      Left            =   5640
      TabIndex        =   4
      Top             =   5640
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
      Left            =   5640
      TabIndex        =   3
      Top             =   5160
      Width           =   2025
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
      Height          =   930
      Left            =   240
      TabIndex        =   0
      Top             =   5040
      Width           =   5235
      Begin VB.CommandButton CmdEliminar 
         Height          =   495
         Left            =   1800
         Picture         =   "FrmMantCompromisoMarcacionDelete.frx":35C4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Confirmar operación"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdCancelar 
         Height          =   510
         Left            =   2760
         Picture         =   "FrmMantCompromisoMarcacionDelete.frx":38CE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cancelar la operación"
         Top             =   240
         Width           =   810
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGFaltas 
      Height          =   3930
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6932
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Idtrabajador"
      Columns(0).DataField=   "Idtrabajador"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NombreCompleto"
      Columns(1).DataField=   "NombreCompleto"
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
      Columns(2).Caption=   "Eliminar"
      Columns(2).DataField=   "Eliminar"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=7144"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=7038"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1826"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1720"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
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
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
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
Attribute VB_Name = "FrmMantCompromisoMarcacionDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
'Public datFecha As Date
Public datFechaInicio As Date
'Public datFechaFin As Date
'Public strRemunerada As String
'Public strObservacion As String
'Private ArrayEliminadas() As String

'Private ArrayAEliminar() As String
Private ArrayTrabajador() As String

Private intSeleccionados As Integer

Public blnStatus As Boolean
Public strMensajeError As String

Private mrstDetalle As Object

Private Const mstrformname As String = "Mantenimiento de Eliminar Marcacion por Grupo"

Public Property Set rstDetalle(ByVal vNewValue As Variant)
    Set mrstDetalle = vNewValue
End Property
Public Property Get rstDetalle() As Variant
    Set rstDetalle = mrstDetalle
End Property



Public Sub ShowForm()
On Error GoTo ErrHandler
    Select Case intAction
'        Case enumActions.Update
'            LoadData
'            SettingControls enumActions.Update, Me
'            Me.Caption = mstrformname & " - " & KEY_ACTION_UPDATE
'        Case enumActions.AddNew
'            LoadData
'            SettingControls enumActions.AddNew, Me
'            Me.Caption = mstrformname & " - " & KEY_ACTION_ADDNEW
'            'Me.DtcMotivoJustificacion.BoundText = 0
'            'Me.DTPFecha.Value = Date
'            Me.DtpFechaInicio.Value = Date
'            'Me.TxtObservacion.Text = ""
'            'Me.TxtIdJustificacion.Text = 0
'            'Me.OptInasistencia.Value = True
'            'Me.OptNoRemunerada.Value = True
        Case enumActions.Delete ' .AddNewGrupo
            LoadData
            SettingControls enumActions.Delete, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_DELETE_GROUP
            'Me.DtcMotivoJustificacion.BoundText = 0
            'Me.DTPFecha.Value = Date
            Me.DtpFechaInicio.Value = Date
            'Me.DtpFechaFin.Value = Date
            'Me.TxtObservacion.Text = ""
            'Me.TxtIdJustificacion.Text = 0
            'Me.OptInasistencia.Value = True
            'Me.OptNoRemunerada.Value = True
    End Select
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantCompromisoMarcacionDelete_ShowForm, Err.Source & " FrmMantCompromisoMarcacionDelete:ShowForm", Err.Description)
End Sub

Private Sub CmdEliminar_Click()
Dim blnResult As Boolean

On Error GoTo ErrHandler
    If Not ValidateFields Then Exit Sub
    Select Case intAction
        Case enumActions.Delete
            blnResult = Delete()
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
'revisa la consulta
On Error GoTo ErrHandler

    If ValidateFieldsConsultar Then
        FrmMensajeEspera.ShowForm
        Set rstDetalle = Nothing
        Set objPersonal = CreateObject("Personal.clsGeneral")
        Set rstDetalle = objPersonal.GetCampos("Entero As Idtrabajador, Texto As NombreCompleto, Boolean as Eliminar")
                
        Set objPersonal = CreateObject("Personal.clsMovimientos")
'        If OptInasistencia.Value Then
'            Set rstFaltas = objPersonal.GetFaltas(Me.intIdTrabajador, Me.datFechaInicio, Me.datFechaFin)
'        End If
'        If OptTardanzas.Value Then
'            Set rstFaltas = objPersonal.GetTardanzas(Me.intIdTrabajador, Me.datFechaInicio, Me.datFechaFin)
'        End If
'        If OptJornadaIncompleta.Value Then
'            Set rstFaltas = objPersonal.GetJornadasIncompletas(Me.intIdTrabajador, Me.datFechaInicio, Me.datFechaFin)
'        End If
        Set rstConsxEliminar = objPersonal.GetConsHorarioxEliminar(Me.datFechaInicio)
        Set objPersonal = Nothing
        
        If rstConsxEliminar.RecordCount > 0 Then
            rstConsxEliminar.MoveFirst
            Do While Not rstConsxEliminar.EOF
                rstDetalle.AddNew
                rstDetalle!IdTrabajador = rstConsxEliminar!IdTrabajador
                rstDetalle!NombreCompleto = rstConsxEliminar!NombreCompleto
                'rstDetalle!Fecha = rstFaltas!Fecha
                'rstDetalle!HoraIngresoIdeal = FormatDateTime(rstFaltas!HoraIngresoIdeal, vbShortTime)
                'rstDetalle!HoraSalidaIdeal = FormatDateTime(rstFaltas!HoraSalidaIdeal, vbShortTime)
                rstDetalle!Eliminar = False
                rstDetalle.Update
                rstConsxEliminar.MoveNext
            Loop
            Set Me.TDBGFaltas.DataSource = mrstDetalle
            FrmMensajeEspera.Hide
        Else
            If rstConsxEliminar.RecordCount > 0 Then
            Me.TDBGFaltas.MoveFirst  '**rstFaltas.MoveFirst 'Me.TDBGFaltas.MoveFirst
            Do While Not Me.TDBGFaltas.EOF
                Me.TDBGFaltas.Delete  '***rstFaltas.Delete
                TDBGFaltas.MoveFirst
            Loop
            End If
            FrmMensajeEspera.Hide
            MsgBox "No existen incidencias registradas...", vbCritical, mstrformname
        End If
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
            If Me.rstDetalle!Eliminar Then
                Me.rstDetalle!Eliminar = False
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
            If Not Me.rstDetalle!Eliminar Then
                Me.rstDetalle!Eliminar = True
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
'*****
Private Sub LoadData()
Dim objPersonal As Object

On Error GoTo ErrHandler
'    Set objPersonal = CreateObject("Personal.clsPersonal")
'    Set rstMotivoJustificacion = objPersonal.GetMotivoJustificacion()
'    Set objPersonal = Nothing

    Set objPersonal = CreateObject("Personal.clsGeneral")
    Set rstDetalle = objPersonal.GetCampos("Entero As Idtrabajador, Texto As NombreCompleto, Boolean as Eliminar")
    Set objPersonal = Nothing
    
Exit Sub
ErrHandler:
    Set objPersonal = Nothing

    Call ErrorMessage(FrmMantCompromisoMarcacionDelete_LoadData, Err.Source & " FrmMantCompromisoMarcacionDelete:LoadData", Err.Description)
End Sub


Public Function ValidateFieldsConsultar() As Boolean
    ValidateFieldsConsultar = False
    
'    If Me.TxtIdTrabajador.Text = 0 Then
'        MsgBox "Debe seleccionar un trabajador", vbInformation, mstrformname
'        Exit Function
'    End If

    'revisar
    'If Me.ChkFecha.Value = 1 Then
        Me.datFechaInicio = Me.DtpFechaInicio.Value
        'Me.datFechaFin = Me.DtpFechaFin.Value
    'Else
        'MsgBox "Usted Esta queriendo Buscar Todas las Incidencias desde la Instalacion del Sistema Desea Continuar?", vbYesNo, break, break
        'Me.datFechaInicio = gdatFecha_Nula_RangoInicio
        'Me.datFechaFin = gdatFecha_Nula_RangoFin
    'End If
    '
    'Me.intIdTrabajador = Val(Me.TxtIdTrabajador.Text)
    ValidateFieldsConsultar = True
    
End Function

Public Function ValidateFields() As Boolean
    ValidateFields = False
    
'    If Trim$(Me.DtcMotivoJustificacion.Text) = "" Then
'        MsgBox "Debe seleccionar el motivo de justificacion", vbInformation, mstrformname
'        Exit Function
'    End If
'
    If Me.rstDetalle.RecordCount <= 0 Then
        MsgBox "Debe seleccionar marcaciones", vbInformation, mstrformname
        Exit Function
    End If
    intSeleccionados = 0
    rstDetalle.MoveFirst
    Do While Not rstDetalle.EOF
        If rstDetalle!Eliminar Then
            intSeleccionados = intSeleccionados + 1
        End If
        rstDetalle.MoveNext
    Loop
    If intSeleccionados = 0 Then
        MsgBox "Debe seleccionar marcaciones para Eliminar", vbInformation, mstrformname
        Exit Function
    Else
        If Not ValidateConsolidado Then
            MsgBox "Existe una Fecha Imposible de Justificar, es posible que se haya consolidado", vbInformation, mstrformname
            Exit Function
        End If
        
        If Not ConfiguraArray Then
            MsgBox "Debe seleccionar marcaciones para justificar", vbInformation, mstrformname
            Exit Function
        End If
        
'        If Me.OptRemunerada.Value Then
'            Me.strRemunerada = gstrHoraRemunerada
'        Else
'            Me.strRemunerada = gstrHoraNoRemunerada
'        End If
'        'Me.intIdTrabajador = 0 'Val(Me.TxtIdTrabajador.Text)
'        Me.intIdMotivoJustificacion = Me.DtcMotivoJustificacion.BoundText
'        Me.strObservacion = Replace(Me.TxtObservacion.Text, "'", "''")
'        Me.datFecha = Me.DTPFecha.Value
        
        ValidateFields = True
    End If
End Function


Public Function Delete() As Boolean
Dim objPersonal As Object
Dim intIdTrabajador As Long
Dim blnResult As Boolean
Dim strMensajeError As String
Dim i As Integer

On Error GoTo ErrHandler
    If Me.TDBGFaltas.SelBookmarks.Count < 0 Then
        MsgBox MSG_SELECTION_EMPTY, vbExclamation, MSG_ACTION_DELETE & " - " & mstrformname
        Delete = False
    Else
        'If MsgBox(MSG_QUESTION_DELETE, vbOKCancel + vbQuestion, MSG_ACTION_DELETE & " - " & mstrformname) = 1 Then
            Screen.MousePointer = vbHourglass
            
            'lngIdMarcacion = mrstDetalle("Idtrabajador")
            
    
            
            Set rstDetalle = Nothing
            Set objPersonal = CreateObject("Personal.clsMovimientos")
            objPersonal.Transaccion_Marcacion ArrayTrabajador(), datFechaInicio
'            If rstDetalle.RecordCount > 0 Then
'                MsgBox strMensajeError, vbCritical, mstrformname
'                Delete = False
'            Else
'                Do While Not rstDetalle.EOF
'                Bitacora "Marcacion", KEY_ACTION_DELETE, "Codigo " & rstDetalle!IdMarcacion
'                rstDetalle.MoveNext
'                Loop
'
'
'            End If
            Set objPersonal = Nothing
            Screen.MousePointer = vbDefault
            LoadData
            Delete = True
        'End If
    End If

Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Screen.MousePointer = vbDefault
    Delete = False
    Call ErrorMessage(FrmConsCompromisoMarcacion_Delete, Err.Source & " FrmConsCompromisoMarcacion:Delete", Err.Description)
End Function



Private Function ConfiguraArray() As Boolean
Dim intContador As Long
Dim intAcumulador As Long
Dim i As Long
Dim intIdTrabajador As Integer
Dim objPersonal As Object

On Error GoTo ErrHandler
intContador = 0
intAcumulador = 0

    ReDim ArrayTrabajador(0 To intSeleccionados)
 
    intSeleccionados = 0
    If Not rstDetalle.BOF And Not rstDetalle.EOF Then
        MsgBox "No existen trabajadores en la lista ", vbInformation, mstrformname
    
    Else
        rstDetalle.MoveFirst

    Do While Not rstDetalle.EOF
        If Me.rstDetalle!Eliminar Then

            intIdTrabajador = Me.rstDetalle!IdTrabajador
            ArrayTrabajador(intSeleccionados) = intIdTrabajador
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


Private Sub TDBGfaltas_Click()
If Me.TDBGFaltas.Col = 4 Then
If Not (Me.rstDetalle.EOF And Me.rstDetalle.BOF) Then
    If Me.rstDetalle!Eliminar Then
        Me.rstDetalle!Eliminar = False
    Else
        Me.rstDetalle!Eliminar = True
    End If
    rstDetalle.Update
End If
End If
End Sub

Public Function ValidateConsolidado() As Boolean
Dim objConsolidado As Object
Dim blnResult As Boolean
On Error GoTo ErrHandler
    ValidateConsolidado = False
    Set objConsolidado = CreateObject("Personal.clsMovimientos")
    Me.rstDetalle.MoveFirst
    Do While Not rstDetalle.EOF
        If rstDetalle!Eliminar Then
            blnResult = objConsolidado.GetExisteConsolidado(Date)   'rstDetalle!idtrabajador
            If blnResult Then
                Set objConsolidado = Nothing
                Exit Function
            End If
        End If
    Me.rstDetalle.MoveNext
    Loop
    Set objConsolidado = Nothing
    ValidateConsolidado = True

Exit Function
ErrHandler:
    Set objConsolidado = Nothing
    ValidateConsolidado = False
    Call ErrorMessage(FrmMantCompromisoMarcacionDelete_ValidateConsolidado, Err.Source & " FrmMantCompromisoMarcacionDelete:ValidateConsolidado", Err.Description)
End Function

