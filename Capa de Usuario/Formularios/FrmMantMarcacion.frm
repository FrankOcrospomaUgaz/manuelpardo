VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D95CB779-00CB-4B49-97B9-9F0B61CAB3C1}#4.0#0"; "biokey.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FrmMantMarcacion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "FrmMantMarcacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8145
   Begin ZKFPEngXControl.ZKFPEngX ZKFPEngX1 
      Left            =   7320
      Top             =   120
      EnrollCount     =   3
      SensorIndex     =   0
      Threshold       =   5
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
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   7440
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   8280
      TabIndex        =   6
      Top             =   6480
      Width           =   735
   End
   Begin VB.TextBox TxtTarjeta 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2970
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   7335
      Width           =   2805
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   4695
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   120
      Picture         =   "FrmMantMarcacion.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   45
      Width           =   510
   End
   Begin TrueOleDBGrid70.TDBGrid TdbgMarcacion 
      Height          =   6525
      Left            =   90
      TabIndex        =   5
      Top             =   660
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   11509
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Apellido"
      Columns(0).DataField=   "Apellido"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nombre"
      Columns(1).DataField=   "Nombre"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Ingreso"
      Columns(2).DataField=   "HoraIngresoReal"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Salida"
      Columns(3).DataField=   "HoraSalidaReal"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).Locked=   -1  'True
      Splits(0).AllowFocus=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4392"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4313"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3969"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3889"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2249"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2170"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
   Begin VB.Label LblTarjeta 
      AutoSize        =   -1  'True
      Caption         =   "Tarjeta :"
      Height          =   195
      Left            =   2220
      TabIndex        =   4
      Top             =   7380
      Width           =   585
   End
   Begin VB.Label LblHora 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   5610
      TabIndex        =   2
      Top             =   120
      Width           =   1485
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   4335
   End
End
Attribute VB_Name = "FrmMantMarcacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intAction As Integer
Public intIdHorario As Long
Public intIdTrabajador As Integer
Public intIdMarcacion As Long
Public datFecha As Date
Public datTime As Date
Public datHoraIngreso As Date
Public datHoraSalida As Date
Public strAccion As String

Public blnStatus As Boolean
Public strMensajeError As String

Private mrstDataSource As Object
Private mrstTrabajador As Object

Private Const mstrformname As String = "Marcaciones por Trabajador"

Dim FTempLen As Integer
Dim FingerCount As Long
Dim fpcHandle As Long
Dim FFingerHuellas() As String
Dim FFingerIds() As Double
Dim FMatchType As Integer

Dim sTemp As String

Dim objPersonalHuella As Object

Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Sub cmdInit_Click()
  ZKFPEngX1.SensorIndex = 0
  If ZKFPEngX1.InitEngine = 0 Then
     'MessageBox 0, "Lector de Huellas iniciado correctamente", "Hint", 0
     
'     StatusBar.Caption = "Lector de Huellas iniciado correctamente"
'     TextSensorCount.Text = ZKFPEngX1.SensorCount & ""
'     TextSensorIndex.Text = ZKFPEngX1.SensorIndex & ""
'     TextSensorSN.Text = ZKFPEngX1.SensorSN
'
'     cmdInit.Enabled = False
     FMatchType = 0
     'ESTE ES EL VALOR A CAMBIAR, DEBEMOS AUMENTARLE o DISMINUIR PARA CONTROLAR REGULACION
     ZKFPEngX1.Threshold = 5
  End If
End Sub

Public Sub cmdVerify_Click()
    If ZKFPEngX1.IsRegister Then
        ZKFPEngX1.CancelEnroll
    End If
'    StatusBar.Caption = "Start Verification(1:1)"
    FMatchType = 1
End Sub

Private Sub ZKFPEngX1_OnCapture(ByVal ActionResult As Boolean, ByVal ATemplate As Variant)
    Dim fi As Long, i As Long
    Dim Score As Long, ProcessNum As Long
    Dim RegChanged As Boolean
    Dim FRegTemplateBK As String
    Dim verificaTrabajador As Boolean
    Dim objPersonal As Object
    Screen.MousePointer = vbHourglass
    
    FRegTemplateBK = ""
    verificaTrabajador = False
    sTemp = ZKFPEngX1.GetTemplateAsString()
  
    verificaTrabajador = False
    For i = 0 To FingerCount
        If ZKFPEngX1.VerFingerFromStr(FFingerHuellas(i), sTemp, False, RegChanged) Then
            verificaTrabajador = True
            Exit For
       End If
    Next
    If verificaTrabajador Then
        Me.datTime = Format(Time, "hh:mm")
        Me.intIdTrabajador = FFingerIds(i)
        If ValidateFieldsHuella Then
            
            Set objPersonal = New clsPersonal
            Set rstTrabajador = objPersonal.GetUnicoTrabajador(FFingerIds(i))
            Set objPersonal = Nothing
            If rstTrabajador.RecordCount <= 0 Then
                Me.intIdTrabajador = 0
            Else
                Me.intIdTrabajador = mrstTrabajador!IdTrabajador
            End If
    
            ' MUESTRA DATOS DEL TRABAJADOR
            Me.Timer.Enabled = False
            FrmMantDetalleMarcacion.intIdMarcacion = Me.intIdMarcacion
            FrmMantDetalleMarcacion.datTime = Me.datTime
            FrmMantDetalleMarcacion.TxtTarjeta.Text = mrstTrabajador!Tarjeta
            FrmMantDetalleMarcacion.TxtApellido.Text = mrstTrabajador!Apellido
            FrmMantDetalleMarcacion.txtNombre.Text = mrstTrabajador!Nombre
            FrmMantDetalleMarcacion.strFoto = mrstTrabajador!Foto
            FrmMantDetalleMarcacion.strAccion = Me.strAccion
            FrmMantDetalleMarcacion.LblFecha = Date
        
            FrmMantDetalleMarcacion.ShowForm
            
    '
    '
    '       FrmMantDetalleMarcacion.id = FFingerIds(i)
    '       FrmMantDetalleMarcacion.intIdMarcacion = Me.intIdMarcacion
    '       FrmMantDetalleMarcacion.strAccion = Me.strAccion
    '       FrmMantDetalleMarcacion.ShowForm
        End If
    Else
       FrmMensajeTrabajador.strMensaje = "Por favor vuelva a registrar su Huella"
        
        FrmMensajeTrabajador.Show
     
    
    '///fin
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub ZKFPEngX1_OnImageReceived(AImageValid As Boolean)
  ZKFPEngX1.PrintImageAt hDC, 8000, 5, ZKFPEngX1.ImageWidth, ZKFPEngX1.ImageHeight
End Sub

Public Property Get rstDataSource() As Variant
    Set rstDataSource = mrstDataSource
End Property
Public Property Set rstDataSource(ByVal vNewValue As Variant)
    Set mrstDataSource = vNewValue
    Set Me.TdbgMarcacion.DataSource = mrstDataSource
End Property

Public Property Get rstTrabajador() As Variant
    Set rstTrabajador = mrstTrabajador
End Property
Public Property Set rstTrabajador(ByVal vNewValue As Variant)
    Set mrstTrabajador = vNewValue
End Property

Public Sub ShowForm()
On Error GoTo ErrHandler
    Select Case intAction
        Case enumActions.AddNew
            LoadData
            SettingControls enumActions.AddNew, Me
            Me.Caption = mstrformname & " - " & KEY_ACTION_ADDNEW
            Me.TxtTarjeta.Text = ""
            Me.LblFecha.Caption = CStr(Date)
            Me.LblHora.Caption = Time
            Me.Timer.Enabled = True
            Me.intIdMarcacion = 0
    End Select
    Me.Show
    Me.ZOrder 0
    Me.TxtTarjeta.SetFocus
Exit Sub
ErrHandler:
    Call ErrorMessage(FrmMantMarcacion_ShowForm, Err.Source & " FrmMantMarcacion:ShowForm", Err.Description)
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim objPersonal As Object
Dim FRegTemplateBK As String
On Error GoTo ErrHandler

    Screen.MousePointer = vbHourglass
    Set objPersonal = CreateObject("Personal.clsMovimientos")
    If Not objPersonal.UpdateEstadoSuspendidoToNormal() Then Exit Sub
    Set objPersonal = Nothing
    Me.Left = (MDIPrincipal.Width - Me.Width) \ 2
    Me.Top = 100
    Me.strMensajeError = ""
    
    FingerCount = 0
    fpcHandle = ZKFPEngX1.CreateFPCacheDB
    cmdInit_Click
    
    Set objPersonal = CreateObject("Personal.clsPersonal")
    Me.datTime = Format(Time, "hh:mm")
    Set rstTrabajador = objPersonal.GetTrabajador()
    Set objPersonal = Nothing
    If rstTrabajador.RecordCount > 0 Then
        rstTrabajador.MoveFirst
        Do While Not rstTrabajador.EOF
            FRegTemplateBK = IIf(IsNull(rstTrabajador!Huella), "", rstTrabajador!Huella)
            
            'ZKFPEngX1.AddRegTemplateStrToFPCacheDB fpcHandle, FingerCount, FRegTemplateBK
                ReDim Preserve FFingerIds(FingerCount + 1)
                ReDim Preserve FFingerHuellas(FingerCount + 1)
            FFingerIds(FingerCount) = rstTrabajador!IdTrabajador
            FFingerHuellas(FingerCount) = FRegTemplateBK
            FingerCount = FingerCount + 1
            
            rstTrabajador.MoveNext
        Loop
        
    End If
    
    Screen.MousePointer = vbDefault
    
    
Exit Sub
ErrHandler:
    Set objPersonal = Nothing
    Call ErrorMessage(FrmMantMarcacion_Load, Err.Source & " FrmMantMarcacion:Load", Err.Description)
End Sub

Private Sub LoadData()
Dim objPersonal As Object

On Error GoTo ErrHandler
    Set objPersonal = CreateObject("Personal.clsMovimientos")
    Set rstDataSource = objPersonal.GetMarcacion(, , Date)
    Set objPersonal = Nothing
    If intIdMarcacion <> 0 Then
        Me.rstDataSource.Find "IdMarcacion=" & intIdMarcacion, , , 1
    End If

Exit Sub
ErrHandler:
    Set objPersonal = Nothing

    Call ErrorMessage(FrmMantMarcacion_LoadData, Err.Source & " FrmMantMarcacion:LoadData", Err.Description)
End Sub

Public Function ValidateFields() As Boolean
Dim objPersonal As Object
Dim rstValidacion As Object
Dim strTipoAccion As String

    ValidateFields = False
    If Me.TxtTarjeta.Text = "" Then
        MsgBox MSG_REQUERY_FIELD, vbInformation, mstrformname
        Exit Function
    End If
    
    If rstTrabajador.RecordCount <= 0 Then
        FrmMensajeTrabajador.strMensaje = "El numero de tarjeta no existe o el Trabajador está Suspendido"
        FrmMensajeTrabajador.Show
        Exit Function
    End If
    
    'Valida si el trabajador tiene compromisos definidos para este dia
    Set objPersonal = CreateObject("Personal.clsMovimientos")
    Set rstValidacion = objPersonal.GetMarcacion(Me.intIdTrabajador, , Date)
    Set objPersonal = Nothing
    If rstValidacion.RecordCount <= 0 Then
        FrmMensajeTrabajador.strMensaje = "No existen compromisos definidos para este día"
        FrmMensajeTrabajador.Show
        Exit Function
    End If
    Set rstValidacion = Nothing
    
    'Determina el codigo de marcacion y devuelve el tipo Entrada o Salida
    Set objPersonal = CreateObject("Personal.clsMovimientos")
    Me.intIdMarcacion = objPersonal.GetMarcacionVarios(Me.intIdTrabajador, Date, Me.datTime, strTipoAccion)
    Set objPersonal = Nothing
    
    If Me.intIdMarcacion = 0 Then
        FrmMensajeTrabajador.strMensaje = "El trabajador ha registrado todas sus marcaciones o el horario no corresponde"
        FrmMensajeTrabajador.Show
        Exit Function
    End If
    
    Me.intIdTrabajador = rstTrabajador!IdTrabajador
    Me.strAccion = strTipoAccion
    ValidateFields = True
End Function

Public Function ValidateFieldsHuella() As Boolean
Dim objPersonal As Object
Dim rstValidacion As Object
Dim strTipoAccion As String

    ValidateFieldsHuella = False
    If rstTrabajador.RecordCount <= 0 Then
        FrmMensajeTrabajador.strMensaje = "El numero de tarjeta no existe o el Trabajador está Suspendido"
        FrmMensajeTrabajador.Show
        Exit Function
    End If
    
    'Valida si el trabajador tiene compromisos definidos para este dia
    Set objPersonal = New clsPersonal
    Set rstValidacion = objPersonal.GetMarcacion(Me.intIdTrabajador, , Date)
    Set objPersonal = Nothing
    If rstValidacion.RecordCount <= 0 Then
        FrmMensajeTrabajador.strMensaje = "No existen compromisos definidos para este día"
        FrmMensajeTrabajador.Show
        Exit Function
    End If
    Set rstValidacion = Nothing
    
    'Determina el codigo de marcacion y devuelve el tipo Entrada o Salida
    Set objPersonal = New clsPersonal
    Me.intIdMarcacion = objPersonal.GetMarcacionVarios(Me.intIdTrabajador, Date, Me.datTime, strTipoAccion)
    Set objPersonal = Nothing
    
    If Me.intIdMarcacion = 0 Then
        FrmMensajeTrabajador.strMensaje = "El trabajador ha registrado todas sus marcaciones o el horario no corresponde"
        FrmMensajeTrabajador.Show
        Exit Function
    End If
    
    'Me.intIdTrabajador = rstTrabajador!IdTrabajador
    Me.strAccion = strTipoAccion
    ValidateFieldsHuella = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    Me.Timer.Enabled = False
    Set mrstDataSource = Nothing
    Set mrstTrabajador = Nothing
    ZKFPEngX1.FreeFPCacheDB (fpcHandle)
End Sub

Private Sub Timer_Timer()
    Me.LblHora.Caption = Time
End Sub

Public Sub TxtTarjeta_KeyPress(KeyAscii As Integer)
Dim FRegTemplateBK As String
Dim fi As Long, i As Long
Dim Score As Long, ProcessNum As Long
Dim RegChanged As Boolean
Dim verificaTrabajador As Boolean
Dim objPersonal As Object

On Error GoTo ErrHandler
    'Funciona para digitar el numero de tarjeta
    If KeyAscii = 13 Then
        Me.datTime = Format(Time, "hh:mm")
        Set objPersonal = CreateObject("Personal.clsPersonal")
        Set rstTrabajador = objPersonal.GetTrabajador(, Me.TxtTarjeta.Text)
        Set objPersonal = Nothing
        If rstTrabajador.RecordCount <= 0 Then
            Me.intIdTrabajador = 0
        Else
            Me.intIdTrabajador = mrstTrabajador!IdTrabajador
        End If
        
        If mrstTrabajador!IndTarjeta <> "S" Then
            Me.TxtTarjeta.Text = ""
            FrmMensajeTrabajador.strMensaje = "No tiene permiso para registra Marcación con su Tarjeta"
            FrmMensajeTrabajador.Show
            Exit Sub
        End If

        If ValidateFields Then

            ' MUESTRA DATOS DEL TRABAJADOR
            Me.Timer.Enabled = False
            FrmMantDetalleMarcacion.intIdMarcacion = Me.intIdMarcacion
            FrmMantDetalleMarcacion.datTime = Me.datTime
            FrmMantDetalleMarcacion.TxtTarjeta.Text = mrstTrabajador!Tarjeta
            FrmMantDetalleMarcacion.TxtApellido.Text = mrstTrabajador!Apellido
            FrmMantDetalleMarcacion.txtNombre.Text = mrstTrabajador!Nombre
            FrmMantDetalleMarcacion.strFoto = mrstTrabajador!Foto
            FrmMantDetalleMarcacion.strAccion = Me.strAccion
            FrmMantDetalleMarcacion.LblFecha = Date

            FrmMantDetalleMarcacion.ShowForm

            Screen.MousePointer = vbDefault
        End If
        Me.TxtTarjeta.Text = ""
    ElseIf KeyAscii = vbKeyF5 Then
            Set objPersonal = CreateObject("Personal.clsMovimientos")
            Set rstDataSource = objPersonal.GetMarcacion(, , Date)
            Set objPersonal = Nothing
        'End If
    End If

Exit Sub
ErrHandler:
    Set objPersonal = Nothing

    Call ErrorMessage(FrmMantMarcacion_AddNew, Err.Source & " FrmMantMarcacion:AddNew", Err.Description)
End Sub
