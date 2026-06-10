VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const mstrClassModule As String = "Personal - cReportes Class Module"

Public Function GetReportTrabajadores(Optional ByVal iintIdTrabajador As Integer = 0, _
                                      Optional ByVal istrEstado As String = "T", _
                                      Optional ByVal iintCodGrupo As Integer = 0, _
                                      Optional ByVal iintCodCategoria As Integer = 0, _
                                      Optional ByVal iintCodCargo As Integer = 0, _
                                      Optional ByVal iintCodArea As Integer = 0) As Object

Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteTrabajadores " & iintIdTrabajador & ",'" & istrEstado & "'," & iintCodCategoria & "," & iintCodGrupo & "," & iintCodCargo & "," & iintCodArea & ""
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetReportTrabajadores = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetReporteTrabajadores, Source:=mstrClassModule & " - GetReportTrabajadores", Description:=Err.Description
End Function

Public Function GetReportHorariosTrab(Optional ByVal iintIdTrabajador As Integer = 0, _
                                    Optional ByVal iintIdGrupo As Integer = 0, _
                                    Optional ByVal iintIdCargo As Integer = 0, _
                                    Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                    Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler
    strSQL = "up_ReporteHorariosTrab " & iintIdTrabajador & "," & iintIdGrupo & "," & iintIdCargo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetReportHorariosTrab = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetReporteHorTrabajadores, Source:=mstrClassModule & " - GetReportHorTrab", Description:=Err.Description
End Function

Public Function GetReportSuspensionesTrab(Optional ByVal iintIdTrabajador As Integer = 0, _
                                    Optional ByVal iintIdGrupo As Integer = 0, _
                                    Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                    Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler
    strSQL = "up_ReporteSuspensionesTrab " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetReportSuspensionesTrab = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetReporteSuspTrabajadores, Source:=mstrClassModule & " - GetReportSuspensionesTrab", Description:=Err.Description
End Function


Public Function GetJustificacionTrab(Optional ByVal iintIdTrabajador As Integer = 0, _
                                    Optional ByVal iintIdGrupo As Integer = 0, _
                                    Optional ByVal iintIdMotivoJustificacion As Integer = 0, _
                                    Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                    Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler
    strSQL = "up_ReporteJustificacionTrab " & iintIdTrabajador & "," & iintIdGrupo & "," & iintIdMotivoJustificacion & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetJustificacionTrab = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetReporteJustificacionTrabajadores, Source:=mstrClassModule & " - GetJustificacionTrab", Description:=Err.Description
    
End Function

Public Function GetJustificacionTrabL(Optional ByVal iintIdTrabajador As Integer = 0, _
                                    Optional ByVal iintIdGrupo As Integer = 0, _
                                    Optional ByVal iintIdMotivoJustificacion As Integer = 0, _
                                    Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                    Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler
    strSQL = "up_ReporteJustificacionTrabL " & iintIdTrabajador & "," & iintIdGrupo & "," & iintIdMotivoJustificacion & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetJustificacionTrabL = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetReporteJustificacionTrabL, Source:=mstrClassModule & " - GetJustificacionTrabL", Description:=Err.Description
    
End Function

Public Function GetRegistroNoMarcacionesD(Optional ByVal iintIdTrabajador As Integer = 0, _
                                        Optional ByVal iintIdGrupo As Integer = 0, _
                                        Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                        Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteRegistroNoMarcacionesD " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRegistroNoMarcacionesD = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRegistroNoMarcacionesD, Source:=mstrClassModule & " - GetRegistroNoMarcacionesD", Description:=Err.Description
End Function

Public Function GetRegistroNoMarcacionesL(Optional ByVal iintIdTrabajador As Integer = 0, _
                                        Optional ByVal iintIdGrupo As Integer = 0, _
                                        Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                        Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteRegistroNoMarcacionesL " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRegistroNoMarcacionesL = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRegistroNoMarcacionesL, Source:=mstrClassModule & " - GetRegistroNoMarcacionesL", Description:=Err.Description
End Function

Public Function GetRegistroNoMarcacionesAmbasD(Optional ByVal iintIdTrabajador As Integer = 0, _
                                        Optional ByVal iintIdGrupo As Integer = 0, _
                                        Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                        Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteRegistroNoMarcacionesAmbasD " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRegistroNoMarcacionesAmbasD = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRegistroNoMarcacionesAmbasD, Source:=mstrClassModule & " - GetRegistroNoMarcacionesAmbasD", Description:=Err.Description
End Function

Public Function GetRegistroNoMarcacionesAmbasL(Optional ByVal iintIdTrabajador As Integer = 0, _
                                        Optional ByVal iintIdGrupo As Integer = 0, _
                                        Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                        Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteRegistroNoMarcacionesAmbasL " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRegistroNoMarcacionesAmbasL = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRegistroNoMarcacionesAmbasL, Source:=mstrClassModule & " - GetRegistroNoMarcacionesAmbasL", Description:=Err.Description
End Function

Public Function GetRegistroNoMarcacionesESD(Optional ByVal iintIdTrabajador As Integer = 0, _
                                        Optional ByVal iintIdGrupo As Integer = 0, _
                                        Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                        Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteRegistroNoMarcacionesESD " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRegistroNoMarcacionesESD = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRegistroNoMarcacionesESD, Source:=mstrClassModule & " - GetRegistroNoMarcacionesESD", Description:=Err.Description
End Function

Public Function GetRegistroNoMarcacionesESL(Optional ByVal iintIdTrabajador As Integer = 0, _
                                        Optional ByVal iintIdGrupo As Integer = 0, _
                                        Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                        Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteRegistroNoMarcacionesESL " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRegistroNoMarcacionesESL = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRegistroNoMarcacionesESL, Source:=mstrClassModule & " - GetRegistroNoMarcacionesESL", Description:=Err.Description
End Function

Public Function GetRecordTardanzasD(Optional ByVal iintIdTrabajador As Integer = 0, _
                                   Optional ByVal iintIdGrupo As Integer = 0, _
                                   Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                   Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String
Dim dblHolgura As Double

On Error GoTo ErrHandler

    dblHolgura = GetToleranciaTardanza()
    
    strSQL = "up_ReporteRecordTardanzasD " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'," & dblHolgura & ""
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRecordTardanzasD = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing

Err.Raise Number:=vbObjectError + clsReportes_GetRecordTardanzasD, Source:=mstrClassModule & " - GetRecordTardanzasD", Description:=Err.Description
End Function

Public Function GetRecordTardanzasL(Optional ByVal iintIdTrabajador As Integer = 0, _
                                   Optional ByVal iintIdGrupo As Integer = 0, _
                                   Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                   Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String
Dim dblHolgura As Double

On Error GoTo ErrHandler
    
    dblHolgura = GetToleranciaTardanza()
        
    strSQL = "up_ReporteRecordTardanzasL " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'," & dblHolgura & ""
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRecordTardanzasL = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    
Err.Raise Number:=vbObjectError + clsReportes_GetRecordTardanzasL, Source:=mstrClassModule & " - GetRecordTardanzasL", Description:=Err.Description
End Function

Public Function GetRecordInasistenciasD(Optional ByVal iintIdTrabajador As Integer = 0, _
                                    Optional ByVal iintIdGrupo As Integer = 0, _
                                    Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                    Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object

Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteRecordInasistenciasD " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRecordInasistenciasD = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRecordInasistenciasD, Source:=mstrClassModule & " - GetRecordInasistenciasD", Description:=Err.Description
End Function

Public Function GetRecordInasistenciasSoliDiaD(Optional ByVal iintIdTrabajador As Integer = 0, _
                                    Optional ByVal iintIdGrupo As Integer = 0, _
                                    Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                    Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object

Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteRecordInasistenciasSoloDiaD " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRecordInasistenciasSoliDiaD = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRecordInasistenciasSoliDiaD, Source:=mstrClassModule & " - GetRecordInasistenciasSoliDiaD", Description:=Err.Description
End Function

Public Function GetRecordInasistenciasL(Optional ByVal iintIdTrabajador As Integer = 0, _
                                    Optional ByVal iintIdGrupo As Integer = 0, _
                                    Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                    Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteRecordInasistenciasL " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRecordInasistenciasL = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRecordInasistenciasL, Source:=mstrClassModule & " - GetRecordInasistenciasL", Description:=Err.Description
End Function

Public Function GetRecordInasistenciasSoloDiaL(Optional ByVal iintIdTrabajador As Integer = 0, _
                                    Optional ByVal iintIdGrupo As Integer = 0, _
                                    Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                    Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteRecordInasistenciasSoloDiaL " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRecordInasistenciasSoloDiaL = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRecordInasistenciasSoliDiaL, Source:=mstrClassModule & " - clsReportes_GetRecordInasistenciasSoliDiaL", Description:=Err.Description
End Function

Public Function GetReportConsolidados(Optional ByVal istrMes As String = "", _
                                      Optional ByVal istrAño As String = "") As Object

Dim objDataAccess As Object
Dim strSQL As String
On Error GoTo ErrHandler

strSQL = "up_ReporteConsolidados '" & istrMes & "','" & istrAño & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetReportConsolidados = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetReportConsolidados, Source:=mstrClassModule & " - GetReportConsolidados", Description:=Err.Description
End Function

Public Function GetBitacora(Optional ByVal istrUsuario As String = "", _
                            Optional ByVal istrTabla As String = "", _
                            Optional ByVal istrOperacion As String = "", _
                            Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                            Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
Dim strSQL As String
Dim objDataAccess As Object
On Error GoTo ErrHandler
    strSQL = "SELECT * FROM Bitacora WHERE 1=1"
    If istrUsuario <> "" Then
        strSQL = strSQL & " AND Usuario = '" & istrUsuario & "'"
    End If
    If istrTabla <> "" Then
        strSQL = strSQL & " AND Tabla = '" & istrTabla & "'"
    End If
    If istrOperacion <> "" Then
        strSQL = strSQL & " AND Operacion = '" & istrOperacion & "'"
    End If
    If idatFechaInicio <> gdatFecha_Nula_RangoInicio Then
        strSQL = strSQL & " AND Fecha >= '" & idatFechaInicio
        End If
    If idatFechaFin <> gdatFecha_Nula_RangoFin Then
        strSQL = strSQL & "' AND Fecha <= '" & idatFechaFin & "'"
    End If
    
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetBitacora = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set GetBitacora = Nothing
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetBitacora, Source:=mstrClassModule & " - GetBitacora", Description:=Err.Description
End Function

Public Function GetDetalleAsistenciaxTrabajador(ByVal idatFechaInicio As Date, _
                                ByVal idatFechaFin As Date, _
                                Optional ByVal iintIdTrabajador As Integer = 0) As Object
Dim strSQL As String
Dim objDataAccess As Object
Dim dblHolguraEntrada As Integer
Dim dblHolguraSalida As Integer

On Error GoTo ErrHandler

    dblHolguraEntrada = GetToleranciaTardanza()
    dblHolguraSalida = GetToleranciaJornadaIncompleta()
    strSQL = "up_ReporteDetalleAsistenciaxTrabajador '" & idatFechaInicio & "','" & idatFechaFin & "'," _
        & dblHolguraEntrada & "," & dblHolguraSalida & "," & iintIdTrabajador & ""
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetDetalleAsistenciaxTrabajador = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
   
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetDetalleAsistenciaxTrabajador, Source:=mstrClassModule & " - GetDetalleAsistenciaxTrabajador", Description:=Err.Description
End Function

Public Function GetDetalleAsistenciaxTrabajadorJust(ByVal idatFechaInicio As Date, _
                                ByVal idatFechaFin As Date, _
                                Optional ByVal iintIdTrabajador As Integer = 0) As Object
Dim strSQL As String
Dim objDataAccess As Object
Dim dblHolguraEntrada As Integer
Dim dblHolguraSalida As Integer

On Error GoTo ErrHandler

    dblHolguraEntrada = GetToleranciaTardanza()
    dblHolguraSalida = GetToleranciaJornadaIncompleta()
    strSQL = "up_ReporteDetalleAsistenciaxTrabajadorJust '" & idatFechaInicio & "','" & idatFechaFin & "'," _
        & dblHolguraEntrada & "," & dblHolguraSalida & "," & iintIdTrabajador & ""
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetDetalleAsistenciaxTrabajadorJust = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
   
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetDetalleAsistenciaxTrabajadorJust, Source:=mstrClassModule & " - GetDetalleAsistenciaxTrabajadorJust", Description:=Err.Description
End Function

Public Function GetDetalleAsistenciaxTrabajadorL(ByVal idatFechaInicio As Date, _
                                ByVal idatFechaFin As Date, _
                                Optional ByVal iintIdTrabajador As Integer = 0) As Object
Dim strSQL As String
Dim objDataAccess As Object
Dim dblHolguraEntrada As Integer
Dim dblHolguraSalida As Integer

On Error GoTo ErrHandler

    dblHolguraEntrada = GetToleranciaTardanza()
    dblHolguraSalida = GetToleranciaJornadaIncompleta()
    strSQL = "up_ReporteDetalleAsistenciaxTrabajadorL '" & idatFechaInicio & "','" & idatFechaFin & "'," _
        & dblHolguraEntrada & "," & dblHolguraSalida & "," & iintIdTrabajador & ""
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetDetalleAsistenciaxTrabajadorL = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetDetalleAsistenciaxTrabajadorL, Source:=mstrClassModule & " - GetDetalleAsistenciaxTrabajadorL", Description:=Err.Description
End Function

Public Function GetRecordSobretiemposD(Optional ByVal iintIdTrabajador As Integer = 0, _
                                   Optional ByVal iintIdGrupo As Integer = 0, _
                                   Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                   Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler
      
    strSQL = "up_ReporteRecordSobretiemposD " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRecordSobretiemposD = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set GetRecordSobretiemposD = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRecordSobretiemposD, Source:=mstrClassModule & " - GetRecordSobretiemposD", Description:=Err.Description
End Function

Public Function GetRecordSobretiemposL(Optional ByVal iintIdTrabajador As Integer = 0, _
                                   Optional ByVal iintIdGrupo As Integer = 0, _
                                   Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                   Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler
      
    strSQL = "up_ReporteRecordSobretiemposL " & iintIdTrabajador & "," & iintIdGrupo & ",'" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRecordSobretiemposL = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set GetRecordSobretiemposL = Nothing
Err.Raise Number:=vbObjectError + clsReportes_GetRecordSobretiemposL, Source:=mstrClassModule & " - GetRecordSobretiemposL", Description:=Err.Description
End Function

Public Function GetRecordsetDesconectado(ByVal istrCampos As String) As Object
Dim strSQL As String
Dim objDataAccess As Object

On Error GoTo ErrHandler
    strSQL = "SELECT " & istrCampos & " FROM Campos"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRecordsetDesconectado = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection, DataAccessLockOptimistic)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set GetRecordsetDesconectado = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetRecordsetDesconectado, mstrClassModule & " - GetRecordsetDesconectado Method", Err.Description, CStr(Erl))
End Function

Public Function GetToleranciaTardanza() As Double
Dim strSQL As String
Dim objDataAccess As Object
Dim rstDataAccess As Object

On Error GoTo ErrHandler
    
    strSQL = "Select ToleranciaTardanza From Cliente "
    Set objDataAccess = CreateObject("Personal.clsData")
    Set rstDataAccess = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    GetToleranciaTardanza = rstDataAccess!ToleranciaTardanza

Exit Function
ErrHandler:
    Set rstDataAccess = Nothing
    GetToleranciaTardanza = 0
    Call RaiseError(vbObjectError + clsReportes_GetToleranciaTardanza, mstrClassModule & " - GetToleranciaTardanza Method", Err.Description, CStr(Erl))
End Function

Public Function GetToleranciaJornadaIncompleta() As Double
Dim strSQL As String
Dim objDataAccess As Object
Dim rstDataAccess As Object

On Error GoTo ErrHandler
    
    strSQL = "Select ToleranciaJornadaIncompleta From Cliente "
    Set objDataAccess = CreateObject("Personal.clsData")
    Set rstDataAccess = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    GetToleranciaJornadaIncompleta = rstDataAccess!ToleranciaJornadaIncompleta

Exit Function
ErrHandler:
    Set rstDataAccess = Nothing
    GetToleranciaJornadaIncompleta = 0
    Call RaiseError(vbObjectError + clsReportes_GetToleranciaJornadaIncompleta, mstrClassModule & " - GetToleranciaJornadaIncompleta Method", Err.Description, CStr(Erl))
End Function

Public Function GetMotivoJustificacion(ByVal iintIdJustificacion As Integer) As Object
Dim strSQL As String
Dim objDataAccess As Object

On Error GoTo ErrHandler
    
    strSQL = "up_BuscarMotivoJustificacion '" & iintIdJustificacion & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetMotivoJustificacion = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set GetMotivoJustificacion = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_UpdateHoraEntradaYSalidaMarcacion, mstrClassModule & " - UpdateHoraEntradaYSalidaMarcacion Method", Err.Description, CStr(Erl))
End Function

Public Function GetMinutoSuspensionTrabajador(ByVal iintIdTrabajador As Integer, _
                                              ByVal idatFechaInicio As Date, _
                                              ByVal idatFechaFin As Date) As Object
Dim strSQL As String
Dim objDataAccess As Object

On Error GoTo ErrHandler
    strSQL = "up_BuscarMinutoSuspensionTrabajador '" & iintIdTrabajador & "','" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetMinutoSuspensionTrabajador = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set GetMinutoSuspensionTrabajador = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_UpdateHoraEntradaYSalidaMarcacion, mstrClassModule & " - UpdateHoraEntradaYSalidaMarcacion Method", Err.Description, CStr(Erl))
End Function

Public Function GetHolgura(ByVal istrDescripcion As String) As Integer
Dim objDataAccess As Object
Dim strSQL As String
On Error GoTo ErrHandler

    strSQL = "up_BuscarHolgura '" & istrDescripcion & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    GetHolgura = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)!ValorMin
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    GetHolgura = 0
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetHolgura, mstrClassModule & " - Update Holgura Method", Err.Description, CStr(Erl))
End Function

Public Function GetRegistroMarcacionesDia(ByVal idatHoraIngresoInicio As Date, _
                                       ByVal idatHoraIngresoFin As Date, _
                                       ByVal idatFecha As Date, _
                                       Optional ByVal iintIdTrabajador As Integer = 0) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteRegistroMarcacionDia '" & FormatDateTime(idatHoraIngresoInicio, vbShortTime) & "','" & FormatDateTime(idatHoraIngresoFin, vbShortTime) & "','" & idatFecha & "','" & iintIdTrabajador & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRegistroMarcacionesDia = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRegistroMarcacionesDia, Source:=mstrClassModule & " - GetRegistroMarcacionesDia", Description:=Err.Description
End Function

Public Function GetRegistroMarcacionesFecha(ByVal idatHoraIngresoInicio As Date, _
                                        ByVal idatHoraIngresoFin As Date, _
                                        ByVal idatFechaInicio As Date, _
                                        ByVal idatFechaFin As Date, _
                                        Optional ByVal iintIdTrabajador As Integer = 0) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteRegistroMarcacionFecha '" & FormatDateTime(idatHoraIngresoInicio, vbShortTime) & "','" & FormatDateTime(idatHoraIngresoFin, vbShortTime) & "','" & idatFechaInicio & "','" & idatFechaFin & "','" & iintIdTrabajador & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRegistroMarcacionesFecha = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRegistroMarcacionesFecha, Source:=mstrClassModule & " - GetRegistroMarcacionesFecha", Description:=Err.Description
End Function

Public Function GetComparacionHoraTrab(ByVal idatFechaInicio As Date, _
                                        ByVal idatFechaFin As Date, _
                                        Optional ByVal iintIdTrabajador As Integer = 0) As Object

Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteComparacionHoraTrab '" & iintIdTrabajador & "','" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetComparacionHoraTrab = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetComparacionHoraTrab, Source:=mstrClassModule & " - GetComparacionHoraTrab", Description:=Err.Description
End Function

Public Function GetComparacionHoraTrabL(ByVal idatFechaInicio As Date, _
                                        ByVal idatFechaFin As Date, _
                                        Optional ByVal iintIdTrabajador As Integer = 0) As Object

Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteComparacionHoraTrabL '" & iintIdTrabajador & "','" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetComparacionHoraTrabL = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetComparacionHoraTrabL, Source:=mstrClassModule & " - GetComparacionHoraTrabL", Description:=Err.Description
End Function

Public Function GetComparacionHoraTrabAsistencia(ByVal idatFechaInicio As Date, _
                                        ByVal idatFechaFin As Date, _
                                        Optional ByVal iintIdGrupo As Integer = 0, _
                                        Optional ByVal iintIdTrabajador As Integer = 0) As Object
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler
    strSQL = "up_ReporteComparacionHoraTrabAsistencia '" & iintIdTrabajador & "','" & iintIdGrupo & "','" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetComparacionHoraTrabAsistencia = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetComparacionHoraTrabAsistencia, Source:=mstrClassModule & " - GetComparacionHoraTrabAsistencia", Description:=Err.Description
End Function

Public Function GetRegistroMarcacionesDiaTipo(ByVal idatHoraIngresoInicio As Date, _
                                       ByVal idatHoraIngresoFin As Date, _
                                       ByVal idatFecha As Date, _
                                       Optional ByVal iintIdTrabajador As Integer = 0, _
                                       Optional ByVal iintIdTipoHora As Integer = 0) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteRegistroMarcacionDiaTipo '" & FormatDateTime(idatHoraIngresoInicio, vbShortTime) & "','" & FormatDateTime(idatHoraIngresoFin, vbShortTime) & "','" & idatFecha & "','" & iintIdTrabajador & "','" & iintIdTipoHora & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRegistroMarcacionesDiaTipo = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRegistroMarcacionesDiaTipo, Source:=mstrClassModule & " - GetRegistroMarcacionesDiaTipo", Description:=Err.Description
End Function

Public Function GetRegistroMarcacionesFechaTipo(ByVal idatHoraIngresoInicio As Date, _
                                        ByVal idatHoraIngresoFin As Date, _
                                        ByVal idatFechaInicio As Date, _
                                        ByVal idatFechaFin As Date, _
                                        Optional ByVal iintIdTrabajador As Integer = 0, _
                                        Optional ByVal iintIdTipoHora As Integer = 0) As Object
                                      
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "up_ReporteRegistroMarcacionFechaTipo '" & FormatDateTime(idatHoraIngresoInicio, vbShortTime) & "','" & FormatDateTime(idatHoraIngresoFin, vbShortTime) & "','" & idatFechaInicio & "','" & idatFechaFin & "','" & iintIdTrabajador & "','" & iintIdTipoHora & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRegistroMarcacionesFechaTipo = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsReportes_GetRegistroMarcacionesFechaTipo, Source:=mstrClassModule & " - GetRegistroMarcacionesFechaTipo", Description:=Err.Description
End Function
