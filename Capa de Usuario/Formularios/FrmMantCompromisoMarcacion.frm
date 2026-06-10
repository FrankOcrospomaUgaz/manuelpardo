VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmMantCompromisoMarcacion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "FrmMantCompromisoMarcacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   6225
   Begin TrueOleDBGrid70.TDBGrid TdbgTrabajador 
      Height          =   6660
      Left            =   150
      TabIndex        =   4
      Top             =   750
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   11748
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
      Columns(1).Caption=   "Nombre Completo"
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
      Columns(2).Caption=   "Seleccionado"
      Columns(2).DataField=   "Seleccion"
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
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._MinWidth=33"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5027"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4948"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._MinWidth=112965008"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(2)._MinWidth=112797568"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=104,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(42)  =   "Named:id=33:Normal"
      _StyleDefs(43)  =   ":id=33,.parent=0"
      _StyleDefs(44)  =   "Named:id=34:Heading"
      _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   ":id=34,.wraptext=-1"
      _StyleDefs(47)  =   "Named:id=35:Footing"
      _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(49)  =   "Named:id=36:Selected"
      _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=37:Caption"
      _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(53)  =   "Named:id=38:HighlightRow"
      _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=39:EvenRow"
      _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(57)  =   "Named:id=40:OddRow"
      _StyleDefs(58)  =   ":id=40,.parent=33"
      _StyleDefs(59)  =   "Named:id=41:RecordSelector"
      _StyleDefs(60)  =   ":id=41,.parent=34"
      _StyleDefs(61)  =   "Named:id=42:FilterBar"
      _StyleDefs(62)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Generar Compromisos"
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
      Left            =   105
      TabIndex        =   2
      Top             =   7950
      Width           =   6015
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
      Left            =   3135
      TabIndex        =   1
      Top             =   7560
      Width           =   2985
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
      Left            =   90
      TabIndex        =   0
      Top             =   7560
      Width           =   2985
   End
   Begin VB.Label LblComentario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Trabajadores que tienen horario, pero no tienen compromisos de marcación generados."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   150
      TabIndex        =   3
      Top             =   90
      Width           =   5880
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmMantCompromisoMarcacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public intIdCategoria As Integer
Public strDescripcion As String
Public blnStatus As Boolean
Public strMensajeError As String
Public intContadorTrab As Integer

Private Array1() As String
'Private Array2() As String
Private ArrayTrabajador() As String

Private mrstPersonal As Object

Private Const mstrformname As String = "Mantenimiento de Compromiso de Marcación"

Public Property Get rstPersonal() As Variant
    Set rstPersonal = mrstPersonal
End Property
Public Property Set rstPersonal(ByVal vNewValue As Variant)
    Set mrstPersonal = vNewValue
    Set Me.TDBGTrabajador.DataSource = mrstPersonal
End Property

Public Sub ShowForm()
On Error GoTo ErrHandler
    Select Case intAction
        Case enumActions.AddNew
            LoadData
            SettingControls enumActions.AddNew, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_ADDNEW
    End Select
    Me.Show
    Me.ZOrder 0

Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantCompromisoMarcacion_ShowForm, Err.Source & " FrmMantCompromisoMarcacion:ShowForm", Err.Description)
End Sub

Private Sub CmdAceptar_Click()
Dim blnResult As Boolean

On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    If Not ValidateFields Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Select Case intAction
        Case enumActions.AddNew
            blnResult = AddNew()
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
    Screen.MousePointer = vbDefault
    
Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    Call ErrorMessage(FrmMantCompromisoMarcacion_ShowForm, Err.Source & " FrmMantCompromisoMarcacion:ShowForm", Err.Description)
End Sub

Private Sub cmdDeseleccionar_Click()
    If rstPersonal.BOF And rstPersonal.EOF Then
        MsgBox "No existen trabajadores en la lista ", vbInformation, mstrformname
    Else
        rstPersonal.MoveFirst
        Do While Not rstPersonal.EOF
            If Me.rstPersonal!Seleccion Then
                Me.rstPersonal!Seleccion = False
            End If
            rstPersonal.Update
            rstPersonal.MoveNext
        Loop
    End If
End Sub

Private Sub cmdSeleccionar_Click()
    If rstPersonal.BOF And rstPersonal.EOF Then
        MsgBox "No existen trabajadores en la lista ", vbInformation, mstrformname
    Else
        rstPersonal.MoveFirst
        Do While Not rstPersonal.EOF
            If Not Me.rstPersonal!Seleccion Then
                Me.rstPersonal!Seleccion = True
            End If
            rstPersonal.Update
            rstPersonal.MoveNext
        Loop
        rstPersonal.MoveFirst
    End If
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.strMensajeError = ""
End Sub

Private Sub LoadData()
Dim objPersonal As Object

On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    Set objPersonal = CreateObject("Personal.clsGeneral")
    Set rstPersonal = objPersonal.GetCampos("Entero As IdTrabajador, Texto As NombreCompleto, Boolean As Seleccion ")
    Set objPersonal = Nothing
    FrmMensajeEspera.ShowForm
    LlenarRstPersonal
    FrmMensajeEspera.Hide
    Screen.MousePointer = vbDefault
    
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Screen.MousePointer = vbDefault
    
    Call ErrorMessage(FrmMantCompromisoMarcacion_LoadData, Err.Source & " FrmMantCompromisoMarcacion:LoadData", Err.Description)
End Sub

Private Sub LlenarRstPersonal()
Dim objPersonal As Object
Dim rstDetalle As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsHorario")
    Set rstDetalle = objPersonal.GetTrabajadorSinCompromiso()
    Set objPersonal = Nothing
    
    If rstDetalle.RecordCount > 0 Then
        rstDetalle.MoveFirst
        
        Do While Not rstDetalle.EOF
            rstPersonal.AddNew
            rstPersonal!IdTrabajador = rstDetalle!IdTrabajador
            rstPersonal!NombreCompleto = rstDetalle!Apellido & " " & rstDetalle!Nombre
            rstPersonal!Seleccion = False
            rstPersonal.Update
        rstDetalle.MoveNext
        Loop
    End If

Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Set rstDetalle = Nothing
    Call ErrorMessage(FrmMantCompromisoMarcacion_LlenarRstPersonal, Err.Source & " FrmMantCompromisoMarcacion:LlenarRstPersonal", Err.Description)
End Sub

Private Function AddNew() As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    
    Set objPersonal = CreateObject("Personal.clsHorario")
    AddNew = objPersonal.AddNewCompromisoMarcacionxTrabajador(Array1(), ArrayTrabajador())
    Set objPersonal = Nothing
    If AddNew Then
        Bitacora "Marcacion", KEY_ACTION_ADDNEW, "Compromiso de Marcacion" & CStr(Date)
        AddNew = True
'    Else
'        MsgBox "Algunos de los Trabajadores seleccionados todavia no tienen Horarios Generados o han sido Eliminados", vbInformation, mstrformname
    End If
    
Exit Function
ErrHandler:
    
    Set objPersonal = Nothing
    AddNew = False
    Call ErrorMessage(FrmMantCompromisoMarcacion_AddNew, Err.Source & " FrmMantCompromisoMarcacion:AddNew", Err.Description)
End Function

Public Function ValidateFields() As Boolean

    intContadorTrab = 0
    ValidateFields = False
        
    If rstPersonal.BOF And rstPersonal.EOF Then
        MsgBox "No existen trabajadores en la lista ", vbInformation, mstrformname
        Exit Function
    Else
        rstPersonal.MoveFirst
        Do While Not rstPersonal.EOF
            If Me.rstPersonal!Seleccion Then
                intContadorTrab = intContadorTrab + 1
            End If
            rstPersonal.MoveNext
        Loop
    End If
    
    If intContadorTrab = 0 Then
        MsgBox "Debe seleccionar al menos 01 trabajador", vbInformation, mstrformname
        Exit Function
    End If
    
    If ConfiguraArray Then
        ValidateFields = True
    End If
    
End Function

Private Function ConfiguraArray() As Boolean
Dim intContador As Long
Dim intAcumulador As Long
Dim i As Long
Dim intIdTrabajador As Integer
Dim objPersonal As Object
Dim rstDetalle As Object

On Error GoTo ErrHandler
intContador = 0
intAcumulador = 0

    ReDim ArrayTrabajador(0 To intContadorTrab)
    intContadorTrab = 0
    rstPersonal.MoveFirst
    'Bucle x trabajador para el q se generan los compromisos
    Do While Not rstPersonal.EOF
        If Me.rstPersonal!Seleccion Then
            intIdTrabajador = Me.rstPersonal!IdTrabajador
            ArrayTrabajador(intContadorTrab) = "up_ModificarTrabajadorxCompromiso '" & intIdTrabajador & "','" & gstrCompromisoGenerado & "'"
            intContadorTrab = intContadorTrab + 1
            Set objPersonal = CreateObject("Personal.clsHorario")
            Set rstDetalle = objPersonal.GetHorarioxTrabajadorSinCompromiso(intIdTrabajador, Date)
            Set objPersonal = Nothing
            'ojo validar vacio por inconsistencia de datos (campo compromisogenerado)
            If rstDetalle.RecordCount > 0 Then
                rstDetalle.MoveFirst
                intAcumulador = intAcumulador + rstDetalle.RecordCount
                ReDim Preserve Array1(0 To intAcumulador)
                '***************************ReDim Preserve Array2(0 To intAcumulador)
                'Bucle para los compromisos de marcacion
                For i = intContador To intAcumulador - 1
                    Array1(i) = "up_AgregarCompromiso '" & intIdTrabajador & "', '" & rstDetalle!Fecha & "','" & FormatDateTime(rstDetalle!HoraIngreso, vbShortTime) & _
                                "','" & FormatDateTime(rstDetalle!HoraSalida, vbShortTime) & "','" & rstDetalle!IdHorario & "'"
                    'Array1(i) = "INSERT INTO Marcacion (IdMarcacion, IdTrabajador, Fecha, HoraIngresoIdeal, HoraSalidaIdeal, IdHorario) VALUES ("
                    'Array2(i) = ", " & intIdTrabajador & ", '" & rstDetalle!Fecha & "','" & rstDetalle!HoraIngreso & _
                     '           "','" & rstDetalle!HoraSalida & "'," & rstDetalle!IdHorario & ")"
                    rstDetalle.MoveNext
                Next i
                intContador = intAcumulador
            End If
            Set rstDetalle = Nothing
        End If
        rstPersonal.MoveNext
    Loop
    ConfiguraArray = True

Exit Function
ErrHandler:
    Set objPersonal = Nothing
    Set rstDetalle = Nothing
    ConfiguraArray = False
    
    Call ErrorMessage(FrmMantCompromisoMarcacion_ConfiguraArray, Err.Source & " FrmMantCompromisoMarcacion:ConfiguraArray", Err.Description)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set rstPersonal = Nothing
End Sub

Private Sub TdbgTrabajador_Click()
    If Not Me.rstPersonal.BOF And Not Me.rstPersonal.EOF Then
    If Me.TDBGTrabajador.Col = 2 Then
        If Me.rstPersonal!Seleccion Then
            Me.rstPersonal!Seleccion = False
        Else
            Me.rstPersonal!Seleccion = True
        End If
        rstPersonal.Update
    End If
    End If
End Sub
