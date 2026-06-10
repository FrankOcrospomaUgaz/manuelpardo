VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsolidado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const mstrClassModule As String = "Personal - clsConsolidado Class Module"

Public Function AddNewConsolidado(ByVal istrMes As String, _
                                  ByVal istrAño As String, _
                                  ByVal iDateFechaInicio As Date, _
                                  ByVal iDateFechaFin As Date) As Boolean
Dim objDataAccess As Object
Dim rstPersonal As Object
Dim strSQL As String
Dim dblHolguraEntrada As Integer
Dim dblHolguraSalida As Integer

On Error GoTo ErrHandler

    dblHolguraEntrada = GetToleranciaTardanza()
    dblHolguraSalida = GetToleranciaJornadaIncompleta()
    
'    '#If TRANSACTION_MODE = 1 Then
'    '    Dim ObjContexto As ObjectContext
'    '    Set ObjContexto = GetObjectContext
'    '#End If
    
    strSQL = "up_AgregarNuevoConsolidado '" & istrMes & "','" & Trim(CStr(istrAño)) & "','" & iDateFechaInicio & "','" & iDateFechaFin & "','" & dblHolguraEntrada & "','" & dblHolguraSalida & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    If Not objDataAccess.RunStoredProcedure(strSQL, gStringConnection) Then
        AddNewConsolidado = False
'        '#If TRANSACTION_MODE = 1 Then
'         '   ObjContexto.SetAbort
'        '#End If
        Exit Function
    End If
    Set objDataAccess = Nothing
    AddNewConsolidado = True
'    '#If TRANSACTION_MODE = 1 Then
'    '    ObjContexto.SetComplete
'    '#End If
Exit Function
ExitMethod:
'    '#If TRANSACTION_MODE = 1 Then
'     '   ObjContexto.SetAbort
'    '#End If
    AddNewConsolidado = False
Exit Function
ErrHandler:
'    '#If TRANSACTION_MODE = 1 Then
'     '   ObjContexto.SetAbort
'    '#End If
    AddNewConsolidado = False
    Call RaiseError(vbObjectError + clsConsolidado_AddNewConsolidado, mstrClassModule & " - AddNewConsolidado Method", Err.Description, CStr(Erl))
End Function

Public Function GetTotalHoraExtra(ByVal idatFechaInicio As Date, ByVal idatFechaFin As Date) As Double
Dim strSQL As String
Dim objDataAccess As Object
Dim rstDataAccess As Object
On Error GoTo ErrHandler
'    '#If TRANSACTION_MODE = 1 Then
'    '    Dim ObjContexto As ObjectContext
'    '    Set ObjContexto = GetObjectContext
'    '#End If
    strSQL = " up_BuscarTotalHoraExtra '" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    If IsNull(rstDataAccess!Diferencia) Then
        GetTotalHoraExtra = 0
    Else
        GetTotalHoraExtra = rstDataAccess!Diferencia * 24
    End If
'    '#If TRANSACTION_MODE = 1 Then
'    '    ObjContexto.SetComplete
'    '#End If
Exit Function
ErrHandler:
'    '#If TRANSACTION_MODE = 1 Then
'     '   ObjContexto.SetAbort
'    '#End If
    GetTotalHoraExtra = 0
    Set objDataAccess = Nothing
    Set rstDataAccess = Nothing
    Call RaiseError(vbObjectError + clsConsolidado_GetTotalHoraExtra, mstrClassModule & " - GetTotalHoraExtra Method", Err.Description, CStr(Erl))
End Function

Public Function GetTrabajadorNoEliminado() As Object
Dim strSQL As String
Dim objDataAccess As Object
On Error GoTo ErrHandler
    strSQL = "SELECT IdTrabajador,(Nombre +' '+ Apellido) As Apellidos FROM Trabajador WHERE Estado <> '" & gstrEstadoTrabajadorEliminado & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetTrabajadorNoEliminado = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set GetTrabajadorNoEliminado = Nothing
    Call RaiseError(vbObjectError + clsConsolidado_GetTrabajadorNoEliminado, mstrClassModule & " - GetTrabajadorNoEliminado Method", Err.Description, CStr(Erl))
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

Public Function GetRecordsetDesconectado(ByVal istrCampos As String) As Object
Dim strSQL As String
Dim objDataAccess As Object
On Error GoTo ErrHandler
    strSQL = "SELECT " & istrCampos & " FROM Campos"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRecordsetDesconectado = objDataAccess.GetDataFromSQLSentence(strSQL, , DataAccessLockOptimistic, DataAccessOpenStatic)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Call RaiseError(vbObjectError + clsMovimientos_UpdateHoraEntradaYSalidaMarcacion, mstrClassModule & " - UpdateHoraEntradaYSalidaMarcacion Method", Err.Description, CStr(Erl))
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

Public Function GetMarcacionesTrabajador(ByVal iintIdTrabajador As Integer, _
                                         ByVal idatFechaInicio As Date, _
                                         ByVal idatFechaFin As Date) As Object
Dim strSQL As String
Dim objDataAccess As Object
On Error GoTo ErrHandler
    strSQL = "up_BuscarMarcacionesTrabajador '" & iintIdTrabajador & "','" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetMarcacionesTrabajador = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Set GetMarcacionesTrabajador = Nothing
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

Public Function GetIncidYMinutoTardNoRemunerado(ByVal iintIdTrabajador As Integer, _
                                          ByVal iDateFechaInicio As Date, _
                                          ByVal iDateFechaFin As Date) As Variant
Dim objDataAccess As Object
Dim objHolgura As Object
Dim strSQL As String
Dim Cantidad As Integer
Dim TotalMin As Double
Dim ArrayIncYMin As Variant
Dim rstDataAccess As Object
Dim dblMinHolgura As Double
On Error GoTo ErrHandler

'    '#If TRANSACTION_MODE = 1 Then
'    '    Dim ObjContexto As ObjectContext
'    '    Set ObjContexto = GetObjectContext
'    '#End If
    Set objHolgura = CreateObject("Personal.clsPersonal")
    dblMinHolgura = objHolgura.GetHolgura("Entrada") / 1440
    Set objHolgura = Nothing
    strSQL = "up_BuscarIncidYMinutoTard '" & iintIdTrabajador & "','" & iDateFechaInicio & "','" & iDateFechaFin & "'," & dblMinHolgura & ",'" & gstrHoraNoRemunerada & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    
    Cantidad = rstDataAccess!Total
    If Not IsNull(rstDataAccess!Diferencia) Then TotalMin = rstDataAccess!Diferencia
        
    strSQL = "up_BuscarIncidYMinutoTardSinJust '" & iintIdTrabajador & "','" & iDateFechaInicio & "','" & iDateFechaFin & "'," & dblMinHolgura & ""
    Set rstDataAccess = Nothing
    Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, , adLockOptimistic, adOpenKeyset)
    
    rstDataAccess!Total = Cantidad + rstDataAccess!Total
    If IsNull(rstDataAccess!Diferencia) Then
        rstDataAccess!Diferencia = TotalMin
    Else
        rstDataAccess!Diferencia = TotalMin + CDbl(rstDataAccess!Diferencia)
    End If
    rstDataAccess.Update
    
    ReDim ArrayIncYMin(rstDataAccess.RecordCount, rstDataAccess.Fields.Count)
    
    RecToArray rstDataAccess, ArrayIncYMin
    
    GetIncidYMinutoTardNoRemunerado = ArrayIncYMin
    
    Set objDataAccess = Nothing
    Set rstDataAccess = Nothing
'    '#If TRANSACTION_MODE = 1 Then
'    '    ObjContexto.SetComplete
'    '#End If
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set rstDataAccess = Nothing
'    '#If TRANSACTION_MODE = 1 Then
'     '   ObjContexto.SetAbort
'    '#End If
    Set objDataAccess = Nothing
    Set rstDataAccess = Nothing
    Set GetIncidYMinutoTardNoRemunerado = ArrayIncYMin ' Que se Devuelve?
    Call RaiseError(vbObjectError + clsConsolidado_GetIncidYMinutoTardNoRemunerado, mstrClassModule & " - GetIncidYMinutoTardNoRemunerado Method", Err.Description, CStr(Erl))
End Function


Public Function GetIncidYMinutoTardRemunerado(ByVal iintIdTrabajador As String, _
                                          ByVal iDateFechaInicio As Date, _
                                          ByVal iDateFechaFin As Date) As Variant

Dim objDataAccess As Object
Dim objHolgura As Object
Dim rstDataAccess As Object
Dim strSQL As String
Dim ArrayIncYMin As Variant
Dim dblMinHolgura As Double
On Error GoTo ErrHandler
    
    '#If TRANSACTION_MODE = 1 Then
    '    Dim ObjContexto As ObjectContext
    '    Set ObjContexto = GetObjectContext
    '#End If
    
    Set objHolgura = CreateObject("Personal.clsPersonal")
    dblMinHolgura = objHolgura.GetHolgura("Entrada") / 1440
    Set objHolgura = Nothing
    strSQL = "up_BuscarIncidYMinutoTard '" & iintIdTrabajador & "','" & iDateFechaInicio & "','" & iDateFechaFin & "'," & dblMinHolgura & ",'" & gstrHoraRemunerada & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    
    Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    ReDim ArrayIncYMin(rstDataAccess.RecordCount, rstDataAccess.Fields.Count)
    RecToArray rstDataAccess, ArrayIncYMin
    GetIncidYMinutoTardRemunerado = ArrayIncYMin
    
    '#If TRANSACTION_MODE = 1 Then
    '    ObjContexto.SetComplete
    '#End If
    Set objDataAccess = Nothing
    Set rstDataAccess = Nothing
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set rstDataAccess = Nothing
        '#If TRANSACTION_MODE = 1 Then
        '   ObjContexto.SetAbort
        '#End If
    Set GetIncidYMinutoTardRemunerado = ArrayIncYMin
    Call RaiseError(vbObjectError + clsConsolidado_GetIncidYMinutoTardRemunerado, mstrClassModule & " - GetIncidYMinutoTardRemunerado Method", Err.Description, CStr(Erl))
End Function

Public Function GetIncidYMinutoInasRemunerada(ByVal iintIdTrabajador As String, _
                                          ByVal iDateFechaInicio As Date, _
                                          ByVal iDateFechaFin As Date) As Variant

Dim objDataAccess As Object
Dim strSQL As String
Dim rstDataAccess As Object
Dim ArrayIncYMin As Variant
On Error GoTo ErrHandler

    '#If TRANSACTION_MODE = 1 Then
    '    Dim ObjContexto As ObjectContext
    '    Set ObjContexto = GetObjectContext
    '#End If

    strSQL = "up_BuscarIncidYMinutoIns '" & iintIdTrabajador & "','" & iDateFechaInicio & "','" & iDateFechaFin & "','" & gstrHoraRemunerada & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
           
    ReDim ArrayIncYMin(rstDataAccess.RecordCount, rstDataAccess.Fields.Count)
    
    RecToArray rstDataAccess, ArrayIncYMin
    
    GetIncidYMinutoInasRemunerada = ArrayIncYMin
    
    Set objDataAccess = Nothing
    Set rstDataAccess = Nothing
    
    '#If TRANSACTION_MODE = 1 Then
    '    ObjContexto.SetComplete
    '#End If
    
    Exit Function
ErrHandler:
       Set objDataAccess = Nothing
       Set rstDataAccess = Nothing
       Set GetIncidYMinutoInasRemunerada = ArrayIncYMin
       '#If TRANSACTION_MODE = 1 Then
         '   ObjContexto.SetAbort
       '#End If

       Call RaiseError(vbObjectError + clsConsolidado_GetIncidYMinutoInasRemunerada, mstrClassModule & " - clsConsolidado_GetIncidYMinutoInasRemunerada Method", Err.Description, CStr(Erl))
End Function

Public Function GetIncidYMinutoInasNoRemunerada(ByVal iintIdTrabajador As String, _
                                          ByVal iDateFechaInicio As Date, _
                                          ByVal iDateFechaFin As Date) As Variant

Dim objDataAccess As Object
Dim strSQL As String
Dim Cantidad As Integer
Dim TotalMin As Double
Dim rstDataAccess As Object
Dim ArrayIncYMin As Variant

On Error GoTo ErrHandler


'#If TRANSACTION_MODE = 1 Then
'    Dim ObjContexto As ObjectContext
'    Set ObjContexto = GetObjectContext
'#End If
    strSQL = "up_BuscarIncidYMinutoInsSinJust '" & iintIdTrabajador & "','" & iDateFechaInicio & "','" & iDateFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    
    Cantidad = rstDataAccess!Total
     If Not IsNull(rstDataAccess!Diferencia) Then TotalMin = rstDataAccess!Diferencia
    
    
    strSQL = "up_BuscarIncidYMinutoInsNo '" & iintIdTrabajador & "','" & iDateFechaInicio & "','" & iDateFechaFin & "','" & gstrHoraRemunerada & "'"
    Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, , adLockOptimistic, adOpenKeyset)
    
    rstDataAccess!Total = Cantidad + rstDataAccess!Total
    If IsNull(rstDataAccess!Diferencia) Then
        rstDataAccess!Diferencia = TotalMin
    Else
        rstDataAccess!Diferencia = TotalMin + CDbl(rstDataAccess!Diferencia)
    End If
    rstDataAccess.Update
    
    ReDim ArrayIncYMin(rstDataAccess.RecordCount, rstDataAccess.Fields.Count)
    
    RecToArray rstDataAccess, ArrayIncYMin
    GetIncidYMinutoInasNoRemunerada = ArrayIncYMin
    
    Set rstDataAccess = Nothing
    Set objDataAccess = Nothing
    
    '#If TRANSACTION_MODE = 1 Then
    '    ObjContexto.SetComplete
    '#End If
    
Exit Function
ErrHandler:
        '#If TRANSACTION_MODE = 1 Then
         '   ObjContexto.SetAbort
        '#End If

       Set objDataAccess = Nothing
       Set rstDataAccess = Nothing
       GetIncidYMinutoInasNoRemunerada = ArrayIncYMin
       Call RaiseError(vbObjectError + clsConsolidado_GetIncidYMinutoInasNoRemunerada, mstrClassModule & " - GetIncidYMinutoInasNoRemunerada Method", Err.Description, CStr(Erl))
End Function

Public Function GetIncidYMinutoSalidaAntRemunerada(ByVal iintIdTrabajador As Integer, _
                                          ByVal iDateFechaInicio As Date, _
                                          ByVal iDateFechaFin As Date) As Variant
Dim objDataAccess As Object
Dim objHolgura As Object
Dim strSQL As String
Dim Cantidad As Integer
Dim TotalMin As Double
Dim rstDataAccess As Object
Dim ArrayIncYMin As Variant
Dim dblMinHolgura As Double
On Error GoTo ErrHandler

    '#If TRANSACTION_MODE = 1 Then
    '    Dim ObjContexto As ObjectContext
    '    Set ObjContexto = GetObjectContext
    '#End If

    Set objDataAccess = CreateObject("Personal.clsData")
    Set objHolgura = CreateObject("Personal.clsPersonal")
    dblMinHolgura = objHolgura.GetHolgura("Salida") / 1440
    Set objHolgura = Nothing
    strSQL = "up_BuscarIncidYMinutoSalidaAnt " & iintIdTrabajador & ",'" & iDateFechaInicio & "','" & iDateFechaFin & "','" & gstrFeriadoRemunerado & "'," & dblMinHolgura & ""
    
    Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    ReDim ArrayIncYMin(rstDataAccess.RecordCount, rstDataAccess.Fields.Count)
    RecToArray objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection), ArrayIncYMin
    
    GetIncidYMinutoSalidaAntRemunerada = ArrayIncYMin
    
    Set objDataAccess = Nothing
    Set rstDataAccess = Nothing
    
    '#If TRANSACTION_MODE = 1 Then
    '    ObjContexto.SetComplete
    '#End If
    
Exit Function
ErrHandler:
        '#If TRANSACTION_MODE = 1 Then
         '   ObjContexto.SetAbort
        '#End If
       Set objHolgura = Nothing
       Set objDataAccess = Nothing
       Set rstDataAccess = Nothing
       GetIncidYMinutoSalidaAntRemunerada = ArrayIncYMin
       Call RaiseError(vbObjectError + clsConsolidado_GetIncidYMinutoSalidaAntRemunerada, mstrClassModule & " - DeleteJustificacion Method", Err.Description, CStr(Erl))
End Function

Public Function GetIncidYMinutoSalidaAntNoRemunerada(ByVal iintIdTrabajador As String, _
                                          ByVal iDateFechaInicio As Date, _
                                          ByVal iDateFechaFin As Date) As Variant
'no se usa esta funcion
Dim objDataAccess As Object
Dim strSQL As String
Dim Cantidad As Integer
Dim TotalMin As Double
Dim ArrayIncYMin As Variant
Dim objHolgura As Object
Dim rstDataAccess As Object
Dim dblMinHolgura As Double
On Error GoTo ErrHandler

    '#If TRANSACTION_MODE = 1 Then
    '    Dim ObjContexto As ObjectContext
    '    Set ObjContexto = GetObjectContext
    '#End If
    Set objHolgura = CreateObject("Personal.clsPersonal")
    dblMinHolgura = objHolgura.GetHolgura("Salida") / 1440
    Set objHolgura = Nothing
    'Para Revisar OJO!!!!!!!!!!!
    
    strSQL = "up_BuscarIncidYMinutoSalidaAntSinJust " & iintIdTrabajador & ",'" & iDateFechaInicio & "','" & iDateFechaFin & "'," & dblMinHolgura & ""
    Set objDataAccess = CreateObject("Personal.clsData")
    Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    
    Cantidad = rstDataAccess!Total
    TotalMin = rstDataAccess!Diferencia
    
    strSQL = "up_BuscarIncidYMinutoSalidaAnt " & iintIdTrabajador & ",'" & iDateFechaInicio & "','" & iDateFechaFin & "','" & gstrHoraNoRemunerada & "'," & dblMinHolgura & ""
    Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    
    rstDataAccess!Total = Cantidad + rstDataAccess!Total
    If IsNull(rstDataAccess!Diferencia) Then
        rstDataAccess!Diferencia = TotalMin
    Else
        rstDataAccess!Diferencia = TotalMin + CDbl(rstDataAccess!Diferencia)
    End If
    rstDataAccess.Update
    ReDim ArrayIncYMin(rstDataAccess.RecordCount, rstDataAccess.Fields.Count)
    RecToArray rstDataAccess, ArrayIncYMin
    GetIncidYMinutoSalidaAntNoRemunerada = ArrayIncYMin
    
    Set objDataAccess = Nothing
    Set rstDataAccess = Nothing
    
    '#If TRANSACTION_MODE = 1 Then
    '    ObjContexto.SetComplete
    '#End If
    
Exit Function
ErrHandler:

       Set objDataAccess = Nothing
       Set rstDataAccess = Nothing
        '#If TRANSACTION_MODE = 1 Then
         '   ObjContexto.SetAbort
        '#End If

       GetIncidYMinutoSalidaAntNoRemunerada = ArrayIncYMin
       Call RaiseError(vbObjectError + clsConsolidado_GetIncidYMinutoSalidaAntNoRemunerada, mstrClassModule & " - GetIncidYMinutoSalidaAntNoRemunerada Method", Err.Description, CStr(Erl))
End Function

Public Function SetCodigoConsolidadoEnMarcacion(ByVal iDateFechaInicio As Date, _
                                                ByVal iDateFechaFin As Date, _
                                                ByVal iintCodigoConsolidado As String) As Boolean
Dim strSQL As String
Dim objDataAccess As Object

On Error GoTo ErrHandler

    '#If TRANSACTION_MODE = 1 Then
    '    Dim ObjContexto As ObjectContext
    '    Set ObjContexto = GetObjectContext
    '#End If

        strSQL = "up_ModificarCodigoConsolidadoMarcacion " & iintCodigoConsolidado & ",'" & iDateFechaInicio & "','" & iDateFechaFin & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        If objDataAccess.RunStoredProcedure(strSQL, gStringConnection) Then
            SetCodigoConsolidadoEnMarcacion = True
        Else
            SetCodigoConsolidadoEnMarcacion = False
        End If
        Set objDataAccess = Nothing
       
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
       
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If

    SetCodigoConsolidadoEnMarcacion = False
    Call RaiseError(vbObjectError + clsPersonal_SetCodigoConsolidadoEnMarcacion, mstrClassModule & " - SetCodigoConsolidadoEnMarcacion Method", Err.Description, CStr(Erl))
End Function

Public Function UpdateConsolidadoAsistencia(ByVal iintCodigoConsolidado As Integer, _
                                            ByVal iintTotalIncidenciaTardanza As Integer, _
                                            ByVal idblTotalMinutoTardanza As Double, _
                                            ByVal idblTotalIncidenciaInasistencia As Integer, _
                                            ByVal idblTotalMinutoJornadaIncompleta As Double, _
                                            ByVal idblTotalHoraExtra As Double, _
                                            ByVal idblTotalHoraTrabajo As Double) As Boolean
    Dim objDataAccess As Object
    Dim strSQL As String
    
On Error GoTo ErrHandler
    '#If TRANSACTION_MODE = 1 Then
    '    Dim ObjContexto As ObjectContext
    '    Set ObjContexto = GetObjectContext
    '#End If
    strSQL = "up_ModificarConsolidadoAsistencia " & iintCodigoConsolidado & "," & iintTotalIncidenciaTardanza & "," & idblTotalMinutoTardanza & ", " & idblTotalIncidenciaInasistencia & "," & idblTotalMinutoJornadaIncompleta & "," & idblTotalHoraExtra & "," & idblTotalHoraTrabajo & ""
    Set objDataAccess = CreateObject("Personal.clsData")
    If Not objDataAccess.RunStoredProcedure(strSQL, gStringConnection) Then
        GoTo ExitMethod
    End If
        UpdateConsolidadoAsistencia = True
    Set objDataAccess = Nothing
              
    '#If TRANSACTION_MODE = 1 Then
    '    ObjContexto.SetComplete
    '#End If
    
Exit Function
ExitMethod:
    Set objDataAccess = Nothing
    UpdateConsolidadoAsistencia = False
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    UpdateConsolidadoAsistencia = False
    Call RaiseError(vbObjectError + clsPersonal_UpdateConsolidadoAsistencia, mstrClassModule & " - UpdateConsolidadoAsistencia Method", Err.Description, CStr(Erl))
End Function

Public Function JustificarFeriadosRemunerados(ByVal iDateFechaInicio As Date, _
                                              ByVal iDateFechaFin As Date) As Boolean

Dim objDataAccess As Object
Dim strSQL As String
Dim intCodigo As Integer
On Error GoTo ErrHandler

    '#If TRANSACTION_MODE = 1 Then
    '    Dim ObjContexto As ObjectContext
    '    Set ObjContexto = GetObjectContext
    '#End If
    Set objDataAccess = CreateObject("Personal.clsData")
    strSQL = "up_AgregarJustificarFeriados " & gintMotivoFeriadoRemunerado & " ,'" & Date & "','" & gstrFeriadoRemunerado & "','" & "Día Feriado Remunerado" & "','" & iDateFechaInicio & "','" & iDateFechaFin & "','" & 1 & "'"
    If Not objDataAccess.RunStoredProcedure(strSQL, gStringConnection) Then
        JustificarFeriadosRemunerados = False
        '#If TRANSACTION_MODE = 1 Then
         '   ObjContexto.SetAbort
        '#End If
        Exit Function
    End If
    Set objDataAccess = Nothing
    JustificarFeriadosRemunerados = True
    '#If TRANSACTION_MODE = 1 Then
    '    ObjContexto.SetComplete
    '#End If
Exit Function
ExitMethod:
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    JustificarFeriadosRemunerados = False
Exit Function
ErrHandler:
    JustificarFeriadosRemunerados = False
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_JustificarFeriadosRemunerados, mstrClassModule & " - JustificarFeriadosRemunerados Method", Err.Description, CStr(Erl))
End Function
Public Function JustificarFeriadosNoRemunerados(ByVal iDateFechaInicio As Date, _
                                   ByVal iDateFechaFin As Date) As Boolean

Dim objDataAccess As Object
Dim strSQL As String
On Error GoTo ErrHandler

    '#If TRANSACTION_MODE = 1 Then
    '    Dim ObjContexto As ObjectContext
    '    Set ObjContexto = GetObjectContext
    '#End If
    Set objDataAccess = CreateObject("Personal.clsData")
    strSQL = "up_AgregarJustificarFeriados " & gintMotivoFeriadoNoRemunerado & " ,'" & Date & "','" & gstrFeriadoNoRemunerado & "','" & "Día Feriado No Remunerado" & "','" & iDateFechaInicio & "','" & iDateFechaFin & "','" & 1 & "'"
    If Not objDataAccess.RunStoredProcedure(strSQL, gStringConnection) Then
        JustificarFeriadosNoRemunerados = False
        '#If TRANSACTION_MODE = 1 Then
         '   ObjContexto.SetAbort
        '#End If
        Exit Function
    End If
    Set objDataAccess = Nothing
    JustificarFeriadosNoRemunerados = True
    '#If TRANSACTION_MODE = 1 Then
    '    ObjContexto.SetComplete
    '#End If
Exit Function
ExitMethod:
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    JustificarFeriadosNoRemunerados = False
Exit Function
ErrHandler:
    JustificarFeriadosNoRemunerados = False
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsConsolidado_JustificarFeriadosNoRemunerados, mstrClassModule & " - JustificarFeriadosNoRemunerados Method", Err.Description, CStr(Erl))
End Function

Public Function GetTotalMinSuspension(ByVal iDateFechaInicio As Date, ByVal iDateFechaFin As Date) As Double
    Dim strSQL As String
    Dim objDataAccess As Object
    Dim rstDataAccess As Object
On Error GoTo ErrHandler

'#If TRANSACTION_MODE = 1 Then
    '    Dim ObjContexto As ObjectContext
    '    Set ObjContexto = GetObjectContext
'#End If
    
    strSQL = "up_BuscarTotalMinSuspension '" & iDateFechaInicio & "','" & iDateFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    
    Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    If IsNull(rstDataAccess!Diferencia) Then
        GetTotalMinSuspension = 0
    Else
        GetTotalMinSuspension = rstDataAccess!Diferencia
    End If
    '#If TRANSACTION_MODE = 1 Then
    '    ObjContexto.SetComplete
    '#End If
Exit Function
ErrHandler:
    GetTotalMinSuspension = 0
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetTotalMinSuspension, mstrClassModule & " - GetTotalMinSuspension Method", Err.Description, CStr(Erl))
End Function

Public Function GetTotalHoraTrabajada(ByVal iDateFechaInicio As Date, ByVal iDateFechaFin As Date) As Double
Dim strSQL As String
Dim objDataAccess As Object
Dim rstDataAccess As Object

On Error GoTo ErrHandler

'#If TRANSACTION_MODE = 1 Then
    '    Dim ObjContexto As ObjectContext
    '    Set ObjContexto = GetObjectContext
'#End If
    strSQL = "up_BuscarTotalHoraTrabajada '" & iDateFechaInicio & "','" & iDateFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    
    Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    If IsNull(rstDataAccess!Diferencia) Then
        GetTotalHoraTrabajada = 0
    Else
        GetTotalHoraTrabajada = rstDataAccess!Diferencia
    End If
    '#If TRANSACTION_MODE = 1 Then
    '    ObjContexto.SetComplete
    '#End If
Exit Function
ErrHandler:
    GetTotalHoraTrabajada = 0
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetTotalHoraTrabajada, mstrClassModule & " - GetTotalHoraTrabajada Method", Err.Description, CStr(Erl))
End Function

Public Function GetTotalHoraDebioTrabajar(ByVal iDateFechaInicio As Date, ByVal iDateFechaFin As Date) As Double
Dim strSQL As String
Dim objDataAccess As Object
Dim rstDataAccess As Object

On Error GoTo ErrHandler

'#If TRANSACTION_MODE = 1 Then
    '    Dim ObjContexto As ObjectContext
    '    Set ObjContexto = GetObjectContext
'#End If
    strSQL = "up_BuscarTotalHoraDebioTrabajar '" & iDateFechaInicio & "','" & iDateFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    
    Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    If IsNull(rstDataAccess!Diferencia) Then
        GetTotalHoraDebioTrabajar = 0
    Else
        GetTotalHoraDebioTrabajar = rstDataAccess!Diferencia
    End If
    '#If TRANSACTION_MODE = 1 Then
    '    ObjContexto.SetComplete
    '#End If
Exit Function
ErrHandler:
    GetTotalHoraDebioTrabajar = 0
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetTotalHoraDebioTrabajar, mstrClassModule & " - GetTotalHoraDebioTrabajar Method", Err.Description, CStr(Erl))
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

