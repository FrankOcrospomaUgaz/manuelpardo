Attribute VB_Name = "ModGeneral"
'****** Constantes del sistema
'Public Const gstrHoraRemunerada As String = "S"
'Public Const gstrHoraNoRemunerada As String = "N"
'Public Const gdatFecha_Nula_RangoInicio As Date = 1 / 1 / 1900
'Public Const gdatFecha_Nula_RangoFin As Date = 31 / 31 / 2999
Public Const gstrEstadoTrabajadorSuspendido As String = "S"
Public Const gstrEstadoTrabajadorEliminado As String = "E"
'Public Const gstrSuspensionEstaActiva As String = "A"
'Public Const gstrSuspensionEstaInactiva As String = "I"
Public Const gintMinutosHolgura As Integer = 60 'minutos
Public Const gstrNoMarcoEntrada As String = "NME"
Public Const gstrNoMarcoSalida As String = "NMS"
Public Const gintGrupoxDefecto As Integer = 0
Public Const gintFeriadoRemunerado As Integer = 1
Public Const gintFeriadoNoRemunerado As Integer = 2
Public Const gintHoraNormal As Integer = 1
Public Const gintHoraExtra As Integer = 2
Public Const gdatHora_Nula As Date = "00:00:00"
Public gstrNameClient As String
Public genumActiveWindow As Integer
Public gstrNombreUsuario As String
Public gaParametrosReport(20, 3) As String
'Public Const gstrCompromisoNoGenerado As String = 0
'Public Const gstrCompromisoGenerado As String = 1
'********* Permisos *********
Public Const gintPermiso As Integer = 3

Public gdatFecha_Nula_RangoInicio As Date
Public gdatFecha_Nula_RangoFin As Date
'****** Horas Nulas *******
Public gdatHora_Nula_RangoInicio As Date
Public gdatHora_Nula_RangoFin As Date

Public Enum enumTipoModificacion
    Ninguna = 1
    SoloIngreso = 2
    SoloSalida = 3
    Ambas = 4
End Enum

'***********Constantes Para el Menú
Public Const MENU_MANTENIMIENTO_AREA As Integer = 1
Public Const MENU_MANTENIMIENTO_CARGO As Integer = 2
Public Const MENU_MANTENIMIENTO_CATEGORIA As Integer = 3
Public Const MENU_MANTENIMIENTO_TIPO_HORA As Integer = 4
Public Const MENU_MANTENIMIENTO_MOTIVO_JUSTIFICACION As Integer = 5
Public Const MENU_MANTENIMIENTO_FERIADO As Integer = 6
Public Const MENU_MANTENIMIENTO_GRUPO As Integer = 7
Public Const MENU_MANTENIMIENTO_TRABAJADOR As Integer = 8
Public Const MENU_MANTENIMIENTO_CLIENTE As Integer = 9
Public Const MENU_MANTENIMIENTO_USUARIO As Integer = 10
Public Const MENU_MANTENIMIENTO_ACCESO_USUARIO As Integer = 11
Public Const MENU_MOVIMIENTO_DEFINIR_HORARIO As Integer = 12
Public Const MENU_MOVIMIENTO_JUSTIFICACION As Integer = 13
Public Const MENU_MOVIMIENTO_PERMISO As Integer = 32
Public Const MENU_MOVIMIENTO_SUSPENSION As Integer = 14
Public Const MENU_MOVIMIENTO_GENERAR_COMP_MARC As Integer = 15
Public Const MENU_MOVIMIENTO_MARCACION As Integer = 16
Public Const MENU_MOVIMIENTO_CONSOLIDADO_MARCACION As Integer = 17
Public Const MENU_REPORTE_TRABAJADORES As Integer = 18
Public Const MENU_REPORTE_HORARIOS As Integer = 19
Public Const MENU_REPORTE_MARCACIONES As Integer = 20
Public Const MENU_REPORTE_COMPARACION As Integer = 21
Public Const MENU_REPORTE_SUSPENSIONES As Integer = 22
Public Const MENU_REPORTE_JUSTIFICACIONES As Integer = 23
Public Const MENU_REPORTE_NO_MARCACIONES As Integer = 24
Public Const MENU_REPORTE_TARDANZAS As Integer = 25
Public Const MENU_REPORTE_INASISTENCIAS As Integer = 26
Public Const MENU_RECORD_SOBRETIEMPOS As Integer = 27
Public Const MENU_DETALLE_ASISTENCIA_TRABAJADOR As Integer = 28
Public Const MENU_CONSOLIDADO_MENSUAL_ASISTENCIA As Integer = 29
Public Const MENU_CONSOLIDADO_ASISTENCIA_TRABAJADOR As Integer = 30
Public Const MENU_BITACORA As Integer = 31



'Coleccion de formularios para almacenar los form con propiedad MDIChild False
Public Colect As New Collection

'****** Constantes de acciones con el sistema
Public Const KEY_ACTION_ADDNEW As String = "[Nuevo]"
Public Const KEY_ACTION_UPDATE As String = "[Modificar]"
Public Const KEY_ACTION_DELETE As String = "[Eliminar]"
Public Const KEY_ACTION_QUERY As String = "[Consultar]"
Public Const KEY_ACTION_PRINT As String = "[Imprimir]"
Public Const KEY_ACTION_FIND As String = "[Buscar]"
Public Const KEY_ACTION_GENERATE As String = "[Generar]"
Public Const KEY_ACTION_CLEAR As String = "[Limpiar]"
Public Const KEY_ACTION_SELECCIONAR As String = "[Seleccionar]"
Public Const KEY_ACTION_ACEPTAR As String = "[Aceptar]"
Public Const KEY_ACTION_AGREGAR As String = "[Agregar]"
Public Const KEY_ACTION_QUITAR As String = "[Quitar]"
Public Const KEY_ACTION_SUSPENDER As String = "[Suspender]"
Public Const KEY_ACTION_ADDNEWGROUP As String = "[NuevoGrupo]"
Public Const KEY_ACTION_DELETE_GROUP As String = "[EliminarGrupo]"
Public Const KEY_ACTION_UPDATE_GROUP As String = "[ModificarGrupo]"
Public Const KEY_ACTION_ADDNEWGROUPDIA As String = "[NuevoGrupoDia]"
Public Enum enumSistemas
    Personal = 1
    Planilla = 2
    Seguridad = 3
End Enum

Public strSistema As enumSistemas

Public Enum enumActions
    AddNew = 1
    Update
    Delete
    Query
    Find
    Printer
    Seting
    Read
    AddNewGrupo 'para el FrmMantJustificacionGrupo
    DeleteGrupo 'para el FrmMantHorarioxGrupoDelete
    UpdateGrupo
    AddNewGrupoDia
End Enum


Dim gobjReporte As Object

'**********************************************************************
'Objetivo:  Define si el control se habilita o no para el formulario .
'Entradas:  iintAction es la acción a realizar
'           objform es el formulario actual
'Retorno:   Ninguno
'**********************************************************************

Public Sub SettingControls(ByVal iintAction As enumActions, objform As Form)
Dim objControl As Control

On Error GoTo ErrHandler
    For Each objControl In objform.Controls
        If Trim(objControl.Tag) <> "" Then
            objControl.Enabled = IIf(objControl.Tag - iintAction >= 0, True, False)
        End If
    Next

Exit Sub
ErrHandler:
    Call ErrorMessage(MGeneral_SettingControls, Err.Source & " MGeneral:SettingControls", Err.Description)
End Sub

'**********************************************************************
'Objetivo:  Determina si el formulario se encuentra activo
'Entradas:  istrFormName es la propiedad Name del formulario
'Retorno:   True si el formulario esta activo
'**********************************************************************

Public Function ActiveForm(ByVal istrFormName As String) As Boolean
    Dim i As Integer
    Dim blnActivo As Boolean
    i = 0
    Do
        blnActivo = (Forms(i).Name = istrFormName)
        i = i + 1
    Loop Until i = Forms.Count Or blnActivo
    ActiveForm = blnActivo
End Function

'**********************************************************************
'Objetivo:  Centra el formulario actual respecto del MDI
'Entradas:  ifrmFormulario es el formulario a centrar
'Retorno:   Ninguno
'**********************************************************************

Public Sub CenterForm(ByRef ifrmFormulario As Form)
    ifrmFormulario.Left = (MDIPrincipal.Width - ifrmFormulario.Width) \ 2
    ifrmFormulario.Top = (MDIPrincipal.Height - ifrmFormulario.Height) \ 2
End Sub

'**********************************************************************
'Objetivo:  Consultar si Desea Tener Holgura o no
'Entradas:  txtTolerancia
'Retorno:   Ninguno
'**********************************************************************
'
'Public Function Holgura(ByVal mstrformname As Integer) As Integer
'   Tolerancia = mstrformname
'End Function


'**********************************************************************
'Objetivo:  Inserta un registro en la Bitacora del sistema
'Entradas:  intColIndex es el número de columna ordenar
'           objRecordset es el objeto del tipo recordset a ordenar
'           tdbGrid es la grilla a modificar
'Retorno:   Ninguno
'**********************************************************************

Public Sub Bitacora(ByVal istrTabla As String, _
                ByVal istrAccion As String, _
                Optional ByVal istrRegistro As String = "")
Dim objGeneral As Object
    Set objGeneral = CreateObject("Personal.clsGeneral")
    objGeneral.AddNewBitacora gstrNombreUsuario, istrTabla, istrAccion, istrRegistro
    Set objGeneral = Nothing
End Sub

'**********************************************************************
'Objetivo:  Ordena la grilla según el número de columna.
'Entradas:  intColIndex es el número de columna ordenar
'           objRecordset es el objeto del tipo recordset a ordenar
'           tdbGrid es la grilla a modificar
'Retorno:   Ninguno
'**********************************************************************

Public Sub OrdenaGrilla(ByVal strDataField As String, objRecordset As Object, tdbGrid As Object)
Dim strOrden As String
Static intAccion As Integer
Static strOldCol As String

On Error GoTo OrdenaGrillaErr
    If strOldCol = strDataField Then
        intAccion = 1 - intAccion
    Else
        intAccion = 1
        strOrden = "ASC"
    End If
    strOrden = IIf(intAccion = 1, "ASC", "DESC")
    If objRecordset.RecordCount = 0 Then Exit Sub
    objRecordset.Sort = strDataField & " " & strOrden
    tdbGrid.Refresh
    strOldCol = strDataField

Exit Sub
OrdenaGrillaErr:
    Call ErrorMessage(MGeneral_OrdenaGrilla, Err.Source & " MGeneral:OrdenaGrilla", Err.Description)
End Sub

Public Sub PrintRecordset(ByVal rstImprimir As Object)
Dim intFilas As Integer
Dim intColumnas As Integer
Dim StrLinea As String
Dim strCampo As String
On Error GoTo ManejoError
If rstImprimir.RecordCount > 0 Then
    rstImprimir.MoveFirst
    For intColumnas = 0 To rstImprimir.Fields.Count - 1
        StrLinea = StrLinea & Chr(9) & Chr(9) & rstImprimir.Fields(intColumnas).Name
    Next
    Debug.Print StrLinea
    StrLinea = ""
    Do While Not rstImprimir.EOF
        For intColumnas = 0 To rstImprimir.Fields.Count - 1
            If CStr(rstImprimir.Fields(intColumnas).Value) = "" Then
                strCampo = "(Vacío)"
            Else
                strCampo = rstImprimir.Fields(intColumnas).Value
            End If
            StrLinea = StrLinea & Chr(9) & Chr(9) & strCampo
        Next
        Debug.Print StrLinea
        StrLinea = ""
        rstImprimir.MoveNext
    Loop
End If
Exit Sub
ManejoError:
    StrLinea = StrLinea & Chr(9) & Chr(9) & "(Vacio)"
End Sub

Public Sub ReBindGrid(ByVal rstDataSource As Object, ByRef tdbGrid As Object)
On Error GoTo ErrHandler
    Set tdbGrid.DataSource = rstDataSource
    tdbGrid.ReBind
    Exit Sub
ErrHandler:
    Call ErrorMessage(modGeneral_ReBindGrid, Err.Source & " modGeneral:ReBindGrid", Err.Description)
End Sub

Public Sub HorarioContinuo(ByRef rstDataSource As Object, ByVal intIdTrabajador As Integer, ByVal datFInicio As Date, _
                            ByVal datFFin As Date, ByVal datHInicio As Date, ByVal datHFin As Date)

Dim datFecha As Date
    
    datFecha = datFInicio
    
    Do While datFecha <= datFFin
    
        rstDataSource.AddNew
        rstDataSource!IdTrabajador = intIdTrabajador
        rstDataSource!FECHA = datFecha
        If datFecha = datFInicio Then
            rstDataSource!HoraIngreso = datHInicio
            rstDataSource!HoraSalida = "11:59 PM"
        End If
        If datFecha = datFFin Then
            rstDataSource!HoraIngreso = "12:00 AM"
            rstDataSource!HoraSalida = datHFin
        End If
        If datFecha <> datFInicio And datFecha <> datFFin Then
            rstDataSource!HoraIngreso = "12:00 AM"
            rstDataSource!HoraSalida = "11:59 PM"
        End If
        
        rstDataSource.Update
        
        datFecha = datFecha + 1
    
    Loop
End Sub
'''' falta
Public Sub HorarioContinuoGrupo(ByRef rstDataSource As Object, ByVal intIdTrabajador As Integer, ByVal datFInicio As Date, _
                            ByVal datFFin As Date, ByVal datHInicio As Date, ByVal datHFin As Date)

Dim datFecha As Date
    
    datFecha = datFInicio
    
    Do While datFecha <= datFFin
    
        rstDataSource.AddNew
        rstDataSource!IdTrabajador = intIdTrabajador
        rstDataSource!FECHA = datFecha
        If datFecha = datFInicio Then
            rstDataSource!HoraIngreso = datHInicio
            rstDataSource!HoraSalida = "11:59 PM"
        End If
        If datFecha = datFFin Then
            rstDataSource!HoraIngreso = "12:00 AM"
            rstDataSource!HoraSalida = datHFin
        End If
        If datFecha <> datFInicio And datFecha <> datFFin Then
            rstDataSource!HoraIngreso = "12:00 AM"
            rstDataSource!HoraSalida = "11:59 PM"
        End If
        
        rstDataSource.Update
        
        datFecha = datFecha + 1
    
    Loop
    
End Sub

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
    'FormReporteMultiple.CRViewer1.ShowExportButton = True
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


Function InsertDataRS(ByRef CDOSet As Object, ByVal irsDatos As Object, ByVal intLabelCount As Integer) As Boolean
On Error GoTo InsertDataErr
Dim LabelRows() As Variant
Dim intX As Double ' As Integer
Dim intC As Integer
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
    Exit Function
InsertDataErr:
    InsertDataRS = False
    Call ErrorMessage(MGeneral_InsertDataRS, Err.Source & "MGeneral:InsertDataRS", Err.Description)
End Function

Public Sub PrintArray(ByVal ArrayImprimir As Variant)
Dim intColumnas As Integer
Dim intFilas As Integer
Dim StrFila As String
On Error GoTo ManejoError
For intFilas = LBound(ArrayImprimir, 1) To UBound(ArrayImprimir, 1)
    For intColumnas = LBound(ArrayImprimir, 2) To UBound(ArrayImprimir, 2)
        StrFila = StrFila & Chr(9) & Chr(9) & CStr(ArrayImprimir(intFilas, intColumnas))
    Next
    Debug.Print StrFila
    StrFila = ""
Next
Exit Sub
ManejoError:
StrFila = StrFila & Chr(9) & Chr(9) & "(Vacío)"
Resume Next
End Sub


Function SetStructuredTTX(ByRef CDOSet As Object, ByVal istrFN As String) As Integer
On Error GoTo ErrHandler
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
ErrHandler:
    SetStructuredTTX = 0
    Call ErrorMessage(MGeneral_SetStructuredTTX, Err.Source & " MGeneral:SetStructuredTTX", Err.Description)
End Function
Public Function RecordsetReporteVacio(ByVal irstReport As Object, _
                                    ByRef istrMensaje As String) As Boolean
    If irstReport.RecordCount = 0 Then
        istrMensaje = "No existen datos para el reporte"
        RecordsetReporteVacio = True
    Else
        RecordsetReporteVacio = False
    End If
End Function


Public Sub ConfiguraNombreCliente()
Dim objGeneral As Object
    Set objGeneral = CreateObject("Personal.clsSeguridad")
    gstrNameClient = objGeneral.GetNombreCliente
    Set objGeneral = Nothing
End Sub
