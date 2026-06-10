Attribute VB_Name = "ModReportes"
Option Explicit


'************************************************************************
'* Nombre : ShowMultiReport
'* Descripción : Muestra un Reporte con o sin SubReportes
'************************************************************************
Public Function ShowMultiReport( _
            irsDatos As Object, _
            istrNombreArchivo As String, _
            Optional arrParametros As Variant = Null, _
            Optional istrPath As String = "", _
            Optional istrTitulo As String = "", _
            Optional iblnParametroUsuario As Boolean = True, _
            Optional iblnPreview As Boolean = True, _
            Optional iblnPermitirVacio As Boolean = False, _
            Optional irsDatos1 As Object = Nothing, _
            Optional istrNombreArchivo1 As String = "", _
            Optional irsDatos2 As Object = Nothing, _
            Optional istrNombreArchivo2 As String = "") As Boolean

'On Error GoTo ShowMultiReportErr
    Dim RepApp As New CRAXDRT.Application
    Dim CDOSet As Object
    Dim RepDb As Object
    Dim RepTables As Object
    Dim RepTable As Object
    Dim intLabelCount As Integer
    Dim objSubReporte1 As Object
    Dim RepApp2 As Object
    Dim CDOSet1 As Object
    Dim RepDb1 As Object
    Dim RepTables1 As Object
    Dim RepTable1 As Object
    Dim intLabelCount1 As Integer
    Dim objSubReporte2 As Object
    Dim CDOSet2 As Object
    Dim RepDb2 As Object
    Dim RepTables2 As Object
    Dim RepTable2 As Object
    Dim intLabelCount2 As Integer
    Dim objParametros As Object
    Dim objParametro As Object
    Dim RepApp1 As Object
    Dim strFN As String
    Dim i As Integer

    If istrTitulo = "" Then istrTitulo = istrNombreArchivo
    'If istrPath = "" Then istrPath = gstrPathReportes
    istrPath = App.Path & "\Reportes\"
'    Set RepApp = New CRAXDRT.Application
    Set gobjReporte = RepApp.OpenReport(istrPath & istrNombreArchivo & ".rpt")
    Set CDOSet = CreateObject("CrystalDataObject.CrystalComObject")
    Set CDOSet1 = CreateObject("CrystalDataObject.CrystalComObject")
    Set CDOSet2 = CreateObject("CrystalDataObject.CrystalComObject")
    
    '...Ingresa estructura de datos del archivo ttx a la estructura CDOSet
        strFN = istrPath & istrNombreArchivo & ".ttx"
        intLabelCount = SetStructuredTTX(CDOSet, strFN)
    
    '...Ahora para el detalle 1 a CDOSet1
    If istrNombreArchivo1 <> "" Then
        strFN = istrPath & istrNombreArchivo1 & ".ttx"
        intLabelCount1 = SetStructuredTTX(CDOSet1, strFN)
    End If
    '...Ahora para el detalle 2 a CDOSet2
    If istrNombreArchivo2 <> "" Then
        strFN = istrPath & istrNombreArchivo2 & ".ttx"
        intLabelCount2 = SetStructuredTTX(CDOSet2, strFN)
    End If
    
    '...Ingreso de datos al objeto CDOSet
    If Not (InsertDataRS(CDOSet, irsDatos, intLabelCount) Or iblnPermitirVacio) Then
        Screen.MousePointer = vbDefault
        MsgBox "El reporte no tiene datos.", vbInformation, istrTitulo
    End If
    '...Asociando la data del objeto CDOSet como tabla del reporte
    Set RepDb = gobjReporte.Database
    Set RepTables = RepDb.Tables
    Set RepTable = RepTables(1)
    Call RepTable.SetPrivateData(3, CDOSet)
    'gobjReporte.DiscardSavedData
    
    '...Existe detalle 1 ?
    If istrNombreArchivo1 <> "" Then
        '...Ingreso de datos al detalle 1
        InsertDataRS CDOSet1, irsDatos1, intLabelCount1
        '...Asociando la data del objeto CDOSet1 a la tabla del SubReporte 1
        Set objSubReporte1 = gobjReporte.OpenSubreport(istrNombreArchivo1 & ".rpt")
        Set RepDb1 = objSubReporte1.Database
        Set RepTables1 = RepDb1.Tables
        Set RepTable1 = RepTables1(1)
        Call RepTable1.SetPrivateData(3, CDOSet1)
        objSubReporte1.DiscardSavedData
    End If
    
    '...Existe detalle 2 ?
    If istrNombreArchivo2 <> "" Then
        '...Ingreso de datos al detalle 2
        InsertDataRS CDOSet2, irsDatos2, intLabelCount2
        '...Asociando la data del objeto CDOSet1 a la tabla del SubReporte 2
        Set objSubReporte2 = gobjReporte.OpenSubreport(istrNombreArchivo2 & ".rpt")
        Set RepDb2 = objSubReporte2.Database
        Set RepTables2 = RepDb2.Tables
        Set RepTable2 = RepTables2(1)
        Call RepTable2.SetPrivateData(3, CDOSet2)
        objSubReporte2.DiscardSavedData
    End If
    
    '...Procesando los datos ingresados en el arreglo arrParametros
    Set objParametros = gobjReporte.ParameterFields
    For Each objParametro In objParametros  '...Para cada parámetro del Reporte
        If UCase(objParametro.ParameterFieldName) = "USUARIO" Then
            'objParametro.SetCurrentValue IIf(iblnParametroUsuario, gstrUsuario, "")
            objParametro.SetCurrentValue IIf(iblnParametroUsuario, 0, "")
        End If
        If Not IsNull(arrParametros) Then   'Si se ha pasado un arreglo de parámetros
            For i = LBound(arrParametros, 1) To UBound(arrParametros, 1)
                If objParametro.ParameterFieldName = CStr(arrParametros(i, 1)) Then
                    objParametro.SetCurrentValue CStr(arrParametros(i, 2))
                End If
            Next    '...Buscando el parámetro
        End If
    Next
    If iblnPreview Then '...Se muestra una presentación preliminar
'    Dim obj As FormReporteMultiple
'    Set obj = New FormReporteMultiple
    FormReporteMultiple.CRViewer1.ReportSource = gobjReporte
    FormReporteMultiple.CRViewer1.ViewReport
    FormReporteMultiple.Show
'        gobjReporte.Preview istrTitulo, , , , , , MDIMenu.hWnd
    Else     '...sino imprimir directamente
'        gobjReporte.Preview istrTitulo, , , , , , MDIPrincipal.hWnd
        gobjReporte.PrintOut False, 1, False
    End If
    ShowMultiReport = True
    'Destruimos los objetos
    Set gobjReporte = Nothing
    Set RepApp = Nothing
    Set CDOSet = Nothing
    Set RepDb = Nothing
    Set RepTables = Nothing
    Set RepTable = Nothing
    Set objParametros = Nothing
    Set objParametro = Nothing
    Set RepApp1 = Nothing
    Set CDOSet1 = Nothing
    Set RepDb1 = Nothing
    Set RepTables1 = Nothing
    Set RepTable1 = Nothing
    Set objSubReporte1 = Nothing
    Set RepApp2 = Nothing
    Set CDOSet2 = Nothing
    Set RepDb2 = Nothing
    Set RepTables2 = Nothing
    Set RepTable2 = Nothing
    Set objSubReporte2 = Nothing
'    Exit Function
'ShowMultiReportErr:
'    Screen.MousePointer = vbDefault
'    ShowMultiReport = False
'    Set RepApp = Nothing
'    Set CDOSet = Nothing
'    Set RepDb = Nothing
'    Set RepTables = Nothing
'    Set RepTable = Nothing
'    Set objParametros = Nothing
'    Set objParametro = Nothing
'    Set RepApp1 = Nothing
'    Set CDOSet1 = Nothing
'    Set RepDb1 = Nothing
'    Set RepTables1 = Nothing
'    Set RepTable1 = Nothing
'    Set objSubReporte1 = Nothing
'    Set RepApp2 = Nothing
'    Set CDOSet2 = Nothing
'    Set RepDb2 = Nothing
'    Set RepTables2 = Nothing
'    Set RepTable2 = Nothing
'    Set objSubReporte2 = Nothing
'    If Err.Number = 429 Then 'error del create object del crystal
'        MsgBox "No tiene instalado correctamente el aplicativo Crystal Reports, no puede imprimirse o mostrarse una vista previa del reporte.", vbInformation, "Error de Instalación del Sistema"
'    ElseIf Err.Number = 20545 Then 'error al ser cancelada la impresión por el usuario
'        MsgBox "La impresión ha sido cancelada por el usuario.", vbInformation, "Impresión Cancelada"
'    ElseIf Err.Number = 20507 Then 'error al no encontrar archivo
'        MsgBox "No se ha encontrado un archivo necesario para la impresión.  Verifique que el sistema esté correctamente instalado.", vbInformation, "No se Encontró Archivo Necesario para Impresión"
'    Else
'        MsgBox Err.Number & " " & Err.Description, vbInformation, MSGERROR
'        Call ErrorMessage(MGeneral_ShowMultiReport, Err.Source & " ModReportes:ShowMultiReport", Err.Description)
'    End If
End Function


'*************************** llenar structura de base de datos***********************

'************************************************************************
'* Nombre : SetStructuresTTX
'* Descripción : Configurar la Estructura de Datos de un Reporte
'************************************************************************
Function SetStructuredTTX(ByRef CDOSet As Object, ByVal istrFN As String) As Integer
'On Error GoTo ErrHandler
Dim intFN As Integer
Dim strLine As String
Dim intLabelCount As Integer
    intFN = FreeFile
    Open istrFN For Input As intFN
    Do While Not EOF(intFN)
        Line Input #intFN, strLine
        If Len(strLine) <> 0 And Right(strLine, 2) <> "%%" Then
            CDOSet.AddField Split(strLine, vbTab)(0), vbString
            intLabelCount = intLabelCount + 1
        End If
    Loop
    Close intFN
    SetStructuredTTX = intLabelCount
    Exit Function
'ErrHandler:
'    SetStructuredTTX = 0
'    MsgBox Err.Number & " " & Err.Description, vbInformation, MSGERROR
'    Call ErrorMessage(MGeneral_SetStructuredTTX, Err.Source & " ModReportes:SetStructuredTTX", Err.Description)
End Function


'**********************insertar data***********************************
'************************************************************************
'* Nombre : InsertDataRS
'* Descripción : Enlaza los Datos del Recordset al Reporte
'************************************************************************
Function InsertDataRS(ByRef CDOSet As Object, ByVal irsDatos As Object, ByVal intLabelCount As Integer) As Boolean
'On Error GoTo InsertDataErr
Dim LabelRows() As Variant
Dim intX As Double
Dim intC As Double
' TODO DPB debe ser parametro iblnPermitirVacio
Dim iblnPermitirVacio As Boolean
    If irsDatos.RecordCount > 0 Or iblnPermitirVacio Then
        With irsDatos
            If Not (.EOF And .BOF) Then
                .MoveLast
                '... LabelRows contiene toda la data del recordset irsDatos
                ReDim LabelRows(.RecordCount - 1, intLabelCount - 1)
                .MoveFirst
                For intX = LBound(LabelRows) To UBound(LabelRows)
                    For intC = 0 To .Fields.Count - 1
                        LabelRows(intX, intC) = CStr(IIf(IsNull(.Fields(intC).Value), "", .Fields(intC).Value))
                    Next 'intC
                    .MoveNext
                Next 'intX
                '...Añade las filas de data del arreglo LabelRows al objeto CDOSet
                CDOSet.AddRows LabelRows
            End If
        End With
    End If
    InsertDataRS = True
'    Exit Function
'InsertDataErr:
'    InsertDataRS = False
'    MsgBox Err.Number & " " & Err.Description, vbInformation, MSGERROR
'    Call ErrorMessage(MGeneral_InsertDataRS, Err.Source & "ModReportes:InsertDataRS", Err.Description)
End Function


