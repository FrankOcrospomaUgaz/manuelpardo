VERSION 5.00
Object = "{D95CB779-00CB-4B49-97B9-9F0B61CAB3C1}#4.0#0"; "biokey.ocx"
Begin VB.Form frmHuella 
   Caption         =   "Biokey Demo"
   ClientHeight    =   7230
   ClientLeft      =   420
   ClientTop       =   1050
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   StartUpPosition =   2  'CenterScreen
   Begin ZKFPEngXControl.ZKFPEngX ZKFPEngX1 
      Left            =   120
      Top             =   1320
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
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2640
      TabIndex        =   25
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Frame FrameCommands 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdEnroll 
         Caption         =   "Register"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "Verify(1:1)"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton cmdIdentify 
         Caption         =   "identify(1:N)"
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox TextFingerName 
         Height          =   390
         Left            =   2640
         TabIndex        =   9
         Text            =   "fingerprint1"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CommandButton cmdInit 
         Caption         =   "Initial Sensor"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   5760
         Width           =   1695
      End
      Begin VB.TextBox TextSensorCount 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1560
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TextSensorIndex 
         Height          =   270
         Left            =   3600
         TabIndex        =   6
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TextSensorSN 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1560
         TabIndex        =   5
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         Caption         =   "Image Format"
         Height          =   615
         Left            =   2040
         TabIndex        =   2
         Top             =   1680
         Width           =   1935
         Begin VB.OptionButton OptionBmp 
            Caption         =   "BMP"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton OptionJpg 
            Caption         =   "JPEG"
            Height          =   255
            Left            =   960
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdSaveImage 
         Caption         =   "Save Image"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox TextWrite 
         Height          =   270
         Left            =   1440
         TabIndex        =   19
         Text            =   "ZKSoftware Biokey SDK."
         Top             =   4560
         Width           =   3015
      End
      Begin VB.TextBox TextRead 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1440
         TabIndex        =   20
         Top             =   4080
         Width           =   3015
      End
      Begin VB.CommandButton cmdReadMemory 
         Caption         =   "Read"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdSaveMemory 
         Caption         =   "Write"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Doggle demo"
         Height          =   1215
         Left            =   120
         TabIndex        =   23
         Top             =   3840
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label StatusBar 
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Sensor Cnt"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Current"
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Serial Number"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmHuella"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FTempLen As Integer
Dim FingerCount As Long
Dim fpcHandle As Long
Dim FFingerNames() As String
Dim FMatchType As Integer

Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Sub cmdEnroll_Click()
  If TextFingerName.Text = "" Then
     MessageBox 0, "Please enter ID", "Error", 0
     Exit Sub
  End If
  ZKFPEngX1.BeginEnroll
  StatusBar.Caption = StatusBar.Caption & Chr(13) & "Inicie registro de huella digital"
End Sub

Private Sub cmdIdentify_Click()
    If ZKFPEngX1.IsRegister Then
        ZKFPEngX1.CancelEnroll
    End If
    StatusBar.Caption = "Start Identification(1:N)"
    FMatchType = 2
End Sub

Public Sub cmdInit_Click()
  ZKFPEngX1.SensorIndex = 0
  If ZKFPEngX1.InitEngine = 0 Then
     'MessageBox 0, "Lector de Huellas iniciado correctamente", "Hint", 0
     
     StatusBar.Caption = "Lector de Huellas iniciado correctamente"
     TextSensorCount.Text = ZKFPEngX1.SensorCount & ""
     TextSensorIndex.Text = ZKFPEngX1.SensorIndex & ""
     TextSensorSN.Text = ZKFPEngX1.SensorSN
     
     cmdInit.Enabled = False
     FMatchType = 0
  End If
End Sub

Private Sub cmdReadMemory_Click()
  Dim i, p1 As Long, p2 As Long
  Dim p As Variant
  Dim s As String
  
  p1 = 0
  p2 = 23
  ZKFPEngX1.DongleMemRead p1, p2, p
  
  s = ""
  For i = LBound(p) To UBound(p)
    s = s & Chr(p(i))
  Next
  
  TextRead.Text = s

End Sub

Private Sub cmdSaveMemory_Click()
  Dim p1 As Long, p2 As Long
  Dim p() As Byte
  Dim i As Integer
  Dim s As String
  
  p1 = 0
  p2 = Len(TextWrite.Text)
  s = TextWrite.Text
  ReDim p(Len(s))
  For i = 0 To Len(s) - 1
    p(i) = Asc(Mid(s, i + 1))
  Next
  ZKFPEngX1.DongleMemWrite p1, p2, p
End Sub

Private Sub cmdSaveImage_Click()
    Dim sFileName As String
    sFileName = "C:\Fingerprint"
    If OptionBmp.Value Then
        ZKFPEngX1.SaveBitmap sFileName & ".bmp"
    Else
        ZKFPEngX1.SaveJPG sFileName + ".jpg"
    End If
    MessageBox 0, "File Saved(" + sFileName + ")!", "", 0
End Sub

Private Sub cmdVerify_Click()
    If ZKFPEngX1.IsRegister Then
        ZKFPEngX1.CancelEnroll
    End If
    StatusBar.Caption = "Start Verification(1:1)"
    FMatchType = 1
End Sub

Private Sub Command1_Click()
    FRegTemplate = ""
    Unload Me
End Sub

Private Sub Command2_Click()

    Unload Me
End Sub

Private Sub Form_Load()
    FingerCount = 0
    fpcHandle = ZKFPEngX1.CreateFPCacheDB
    Me.cmdInit_Click
    cmdEnroll_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ZKFPEngX1.FreeFPCacheDB (fpcHandle)
End Sub

Private Sub ZKFPEngX1_OnCapture(ByVal ActionResult As Boolean, ByVal ATemplate As Variant)
    Dim fi As Long, i As Long
    Dim Score As Long, ProcessNum As Long
    Dim RegChanged As Boolean
    Dim sTemp As String
    
    sTemp = ZKFPEngX1.GetTemplateAsString()
  
    StatusBar.Caption = "Acqired Template"
    If FMatchType = 1 Then  '1:1
       If ZKFPEngX1.VerFinger(FRegTemplate, ATemplate, False, RegChanged) Then
          MessageBox 0, "Verificacion correcta", "information", 0
       Else
          MessageBox 0, "Verificacion fallida", "information", 0
       End If
       'If ZKFPEngX1.VerRegFingerFile("c:\fingerprint.tpl", ATemplate) Then
       '   MessageBox 0, "Verify success", "information", 0
       'Else
       '   MessageBox 0, "Verify failed", "information", 0
       'End If
    ElseIf FMatchType = 2 Then '1:N
       Score = 8
       fi = ZKFPEngX1.IdentificationFromStrInFPCacheDB(fpcHandle, sTemp, Score, ProcessNum)
       If fi = -1 Then
          MessageBox 0, "Identificacion fallada", "information", 0
       Else
          MessageBox 0, "Indentificacion correcta : Nombre=" & FFingerNames(fi) & " Score = " & Score & " Processed Number = " & ProcessNum, "information", 0
       End If
    End If
End Sub

Private Sub ZKFPEngX1_OnEnroll(ByVal ActionResult As Boolean, ByVal ATemplate As Variant)
  Dim i As Long
  Dim sTemp As String
  
  If Not ActionResult Then
    MessageBox 0, "Registro fallido", "Warning", 0
    cmdEnroll_Click
  Else
    MessageBox 0, "Registro correcto", "Information", 0
    'FRegTemplate = ATemplate
    
    
    '得到字符串方式的模板
    FRegTemplate = ZKFPEngX1.GetTemplateAsString()
    '可以直接存入数据库
    
    
    'ZKFPEngX1.SaveTemplate "c:\fingerprint.tpl", ATemplate
    'ZKFPEngX1.EncodeTemplate ATemplate, sTemp
    
    'ZKFPEngX1.AddRegTemplateFileToFPCacheDB fpcHandle, FingerCount, "c:\fingerprint.tpl"
    '    ReDim Preserve FFingerNames(FingerCount + 1)
'    ZKFPEngX1.AddRegTemplateStrToFPCacheDB fpcHandle, FingerCount, sTemp
'        ReDim Preserve FFingerNames(FingerCount + 1)
'    FFingerNames(FingerCount) = TextFingerName.Text
'    FingerCount = FingerCount + 1
    'Unload Me
  End If
End Sub

Private Sub ZKFPEngX1_OnFeatureInfo(ByVal AQuality As Long)
  Dim sTemp As String
  Dim i As Integer
  For i = 1 To 10
    Beep
  Next
  sTemp = ""
  If ZKFPEngX1.IsRegister Then
     sTemp = "Estado de Registro: Todavia le falta realizar " & ZKFPEngX1.EnrollIndex - 1 & " veces el Registro de su Huella"
  End If
  sTemp = sTemp & " - Calidad de Huella Digital"
  If AQuality <> 0 Then
     sTemp = sTemp & " Mala " & AQuality
  Else
     sTemp = sTemp & " Buena"
  End If
  StatusBar.Caption = sTemp
  MessageBox 0, sTemp, "Information", 0
End Sub

Private Sub ZKFPEngX1_OnImageReceived(AImageValid As Boolean)
  ZKFPEngX1.PrintImageAt hDC, 27, 80, ZKFPEngX1.ImageWidth, ZKFPEngX1.ImageHeight
End Sub
