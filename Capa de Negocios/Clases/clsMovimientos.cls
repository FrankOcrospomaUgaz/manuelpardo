VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mstrClassModule As String = "Personal - cMovimientos Class Module"

'Agrega registro suspension y modifica el estado del trabajador

Public Function AddNewSuspension(ByVal iintIdTrabajador As Integer, _
                                ByVal idatFechaInicio As Date, _
                                ByVal idatFechaFin As Date, _
                                Optional ByVal istrMotivo As String = "") As Boolean
    
Dim objDataAccess As Object
Dim intCodigo As Long
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_AgregarSuspension '" & iintIdTrabajador & "','" & istrMotivo & "','" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    AddNewSuspension = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewSuspension = False
    Call RaiseError(vbObjectError + clsMovimientos_AddNewSuspension, mstrClassModule & " - AddNewSuspension Method", Err.Description, CStr(Erl))
    
End Function

Public Function UpdateSuspension(ByVal iintIdSuspension As Long, _
                            ByVal idatFechaInicio As Date, _
                            ByVal idatFechaFin As Date, _
                            Optional ByVal istrMotivo As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_ModificarSuspension '" & idatFechaInicio & "','" & idatFechaFin & "','" & istrMotivo & "','" & iintIdSuspension & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    UpdateSuspension = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    UpdateSuspension = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateSuspension = False
    Call RaiseError(vbObjectError + clsMovimientos_UpdateSuspension, mstrClassModule & " - UpdateSuspension Method", Err.Description, CStr(Erl))
    
End Function

'Elimina registro suspension y modifica el estado del trabajador

Public Function DeleteSuspension(ByVal iintIdSuspension As Long, _
                        ByVal iintIdTrabajador As Integer, _
                        ByVal istrEstadoSuspension As String, _
                        Optional ByRef ostrError As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    If ValidateBussines(EnumActionsMovimientos.DeleteSuspension, ostrError, , _
                        iintIdSuspension) Then
        strSQL = "up_EliminarSuspension '" & iintIdSuspension & "','" & iintIdTrabajador & "','" & istrEstadoSuspension & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        DeleteSuspension = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Set objDataAccess = Nothing
    Else
        GoTo ExitMethod
    End If

Exit Function
ExitMethod:
    Set objDataAccess = Nothing
    DeleteSuspension = False
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    DeleteSuspension = False
    Call RaiseError(vbObjectError + clsMovimientos_DeleteSuspension, mstrClassModule & " - DeleteSuspension Method", Err.Description, CStr(Erl))
    
End Function

Public Function GetSuspension(Optional ByVal iintIdSuspension As Long = 0, _
                            Optional ByVal iintIdTrabajador As Integer = 0, _
                            Optional ByVal istrMotivo As String = "", _
                            Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                            Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
  strSQL = "SELECT IdSuspension, Suspension.IdTrabajador, Apellido + ' ' + Nombre AS Apellido, Motivo, FechaInicio, FechaFin " _
                & " FROM Suspension INNER JOIN Trabajador ON Suspension.IdTrabajador = Trabajador.IdTrabajador WHERE 1=1 "
             
    If iintIdSuspension <> 0 Then
      strSQL = strSQL & " AND IdSuspension = " & iintIdSuspension
    End If
    If iintIdTrabajador <> 0 Then
      strSQL = strSQL & " AND IdTrabajador = " & iintIdTrabajador
    End If
    If istrMotivo <> "" Then
      strSQL = strSQL & " AND Motivo = '" & istrMotivo & "'"
    End If
    If idatFechaInicio <> gdatFecha_Nula_RangoInicio And idatFechaFin <> gdatFecha_Nula_RangoFin Then
        strSQL = strSQL & " AND FechaInicio >= '" & idatFechaInicio & "' AND FechaFin <= '" & idatFechaFin & "'"
    End If
    
  Set objDataAccess = CreateObject("Personal.clsData")
  Set GetSuspension = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
  Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetSuspension, mstrClassModule & " - GetSuspension Method", Err.Description, CStr(Erl))

End Function

' modifica registro de marcacion, agregando el ingreso o salida real
Public Function UpdateMarcacion(ByVal iintIdMarcacion As Long, _
                                ByVal idatHoraReal As Date, _
                                ByVal istrTipoMarcacion As String) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    strSQL = "up_ModificarMarcacion '" & iintIdMarcacion & "','" & FormatDateTime(idatHoraReal, vbShortTime) & "','" & FormatDateTime(idatHoraReal, vbShortTime) & "','" & FormatDateTime(idatHoraReal, vbShortTime) & "','" & FormatDateTime(idatHoraReal, vbShortTime) & "','" & istrTipoMarcacion & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    UpdateMarcacion = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateMarcacion = False
    Call RaiseError(vbObjectError + clsMovimientos_UpdateMarcacion, mstrClassModule & " - UpdateMarcacion Method", Err.Description, CStr(Erl))
End Function

Public Function GetMarcacion(Optional ByVal iintIdTrabajador As Integer = 0, _
                        Optional ByVal iintIdJustificacion As Long = 0, _
                        Optional ByVal idatFecha As Date = gdatFecha_Nula, _
                        Optional ByVal istrNoMarcoEntrada As String = "", _
                        Optional ByVal istrNoMarcoSalida As String = "") As Object
                                
Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
    strSQL = "SELECT IdMarcacion, Marcacion.IdTrabajador, Apellido, Nombre, Marcacion.Fecha, dbo.formatdate(Marcacion.HoraIngresoIdeal,'hh:nn tt') as HoraIngresoIdeal, " & _
            " dbo.formatdate(Marcacion.HoraSalidaIdeal,'hh:nn tt') as HoraSalidaIdeal , dbo.formatdate(Marcacion.HoraIngresoReal, 'hh:nn tt') as HoraIngresoReal, dbo.formatdate(Marcacion.HoraSalidaReal,'hh:nn tt') as HoraSalidaReal " & _
            " FROM Marcacion INNER JOIN Trabajador ON Trabajador.IdTrabajador = Marcacion.IdTrabajador " & _
            " WHERE 1=1 "
             
    If iintIdTrabajador <> 0 Then
        strSQL = strSQL & " AND Marcacion.IdTrabajador = " & iintIdTrabajador
    End If
       
    If iintIdJustificacion <> 0 Then
        strSQL = strSQL & " AND Marcacion.IdJustificacion = " & iintIdJustificacion
    End If
    If idatFecha <> gdatFecha_Nula Then
        strSQL = strSQL & " AND Marcacion.Fecha = '" & idatFecha & "'"
        If istrNoMarcoEntrada <> "" Then
            strSQL = strSQL & " AND Marcacion.HoraIngresoReal IS NULL"
        End If
    End If
    
    If idatFecha <> gdatFecha_Nula Then
        strSQL = strSQL & " AND Marcacion.Fecha = '" & idatFecha & "'"
        If istrNoMarcoSalida <> "" Then
            strSQL = strSQL & " AND Marcacion.HoraSalidaReal IS NULL"
        End If
    End If
    strSQL = strSQL & " ORDER BY Trabajador.Apellido"
    
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetMarcacion = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetMarcacion, mstrClassModule & " - GetMarcacion Method", Err.Description, CStr(Erl))

End Function

'considera la consulta para marcaciones de varios compromisos el mismo dia
'devuelve el codigo de marcacion q corresponde y como parametro el tipo de marcacion: Ent o Sal
Public Function GetMarcacionVarios(ByVal iintIdTrabajador As Integer, _
                        ByVal idatFecha As Date, _
                        ByVal idatHoraReal As Date, _
                        ByRef ostrTipoMarcacion As String) As Long
                                
Dim objDataAccess As Object
Dim rstMarcaciones As Object
Dim rstMarcacionesTolerancia As Object
Dim strSQL As String
Dim dblMinutosHolgura As Double
Dim dblMinutosPrimeraHolgura As Double
Dim lngPrimeraMarcacion As Long
Dim lngUltimaMarcacion As Long
Dim datPrimeraHoraIngreso As Date

On Error GoTo ErrHandler
 
    '*** OBTIENE LAS MARCACIONES DEL DIA DEL TRABAJADOR
    strSQL = "SELECT IdMarcacion, Marcacion.Fecha, Marcacion.HoraIngresoIdeal, Marcacion.HoraSalidaIdeal, " & _
            " Marcacion.HoraIngresoReal, Marcacion.HoraSalidaReal FROM Marcacion " & _
            " WHERE Marcacion.IdTrabajador = " & iintIdTrabajador & _
            " AND Fecha = '" & Date & _
            "'AND (Marcacion.HoraIngresoReal is null OR Marcacion.HoraSalidaReal is null) " & _
            "ORDER BY Marcacion.HoraSalidaIdeal"
            
    Set objDataAccess = CreateObject("Personal.clsData")
    Set rstMarcaciones = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    '*** OBTENER MODO DE CAPTURA DE MARCACION Y DE EXISTIR EL TOTAL DE MINUTOS
    strSQL = "SELECT Tolerancia, Tiempo, TiempoPrimeraMarcacion, HoraPrimeraMarcacion from Cliente"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set rstMarcacionesTolerancia = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    
    If rstMarcacionesTolerancia.RecordCount > 0 Then
        If rstMarcacionesTolerancia!Tolerancia Then
            dblMinutosHolgura = rstMarcacionesTolerancia!Tiempo
            dblMinutosPrimeraHolgura = rstMarcacionesTolerancia!TiempoPrimeraMarcacion
            datPrimeraHoraIngreso = FormatDateTime(rstMarcacionesTolerancia!HoraPrimeraMarcacion, vbShortTime)
            
            '=============================
            '**** PROCEDIMIENTO CON HOLGURA
            '=============================
            If rstMarcaciones.RecordCount = 0 Then Exit Function
                rstMarcaciones.MoveFirst
                'almacena en las variables el codigo de la primera
                lngPrimeraMarcacion = rstMarcaciones!idmarcacion
                
                'Por cada compromiso de marcacion para este dia
                Do While Not rstMarcaciones.EOF
                    If IsNull(rstMarcaciones!HoraIngresoReal) Then
                        'la primera hora de entrada del dia
                        If lngPrimeraMarcacion = rstMarcaciones!idmarcacion Then
                            'Es la hora definida como inicial para considerar mayor holgura
                            If datPrimeraHoraIngreso = FormatDateTime(rstMarcaciones!HoraIngresoIdeal, vbShortTime) Then
                                If DateDiff("n", idatHoraReal, FormatDateTime(DateAdd("n", -dblMinutosPrimeraHolgura, rstMarcaciones!HoraIngresoIdeal), vbShortTime)) <= 0 _
                                    And DateDiff("n", idatHoraReal, FormatDateTime(DateAdd("n", dblMinutosHolgura, rstMarcaciones!HoraIngresoIdeal), vbShortTime)) >= 0 Then
                                    GetMarcacionVarios = rstMarcaciones!idmarcacion
                                    ostrTipoMarcacion = gstrEntrada
                                    Exit Do
                                End If
                            Else
                                If DateDiff("n", idatHoraReal, FormatDateTime(DateAdd("n", -dblMinutosHolgura, rstMarcaciones!HoraIngresoIdeal), vbShortTime)) <= 0 _
                                    And DateDiff("n", idatHoraReal, FormatDateTime(DateAdd("n", dblMinutosHolgura, rstMarcaciones!HoraIngresoIdeal), vbShortTime)) >= 0 Then
                                    GetMarcacionVarios = rstMarcaciones!idmarcacion
                                    ostrTipoMarcacion = gstrEntrada
                                    Exit Do
                                End If
                            End If
                        'Cualquier otra hora diferente de la primera
                        Else
                            If DateDiff("n", idatHoraReal, FormatDateTime(DateAdd("n", -dblMinutosHolgura, rstMarcaciones!HoraIngresoIdeal), vbShortTime)) <= 0 _
                                And DateDiff("n", idatHoraReal, FormatDateTime(DateAdd("n", dblMinutosHolgura, rstMarcaciones!HoraIngresoIdeal), vbShortTime)) >= 0 Then
                                GetMarcacionVarios = rstMarcaciones!idmarcacion
                                ostrTipoMarcacion = gstrEntrada
                                Exit Do
                            End If
                        End If
                    End If
                    'Si la salida esta vacia
                    If IsNull(rstMarcaciones!HoraSalidaReal) Then
                        If DateDiff("n", idatHoraReal, FormatDateTime(DateAdd("n", -dblMinutosHolgura, rstMarcaciones!HoraSalidaIdeal), vbShortTime)) <= 0 _
                            And DateDiff("n", idatHoraReal, FormatDateTime(DateAdd("n", dblMinutosHolgura, rstMarcaciones!HoraSalidaIdeal), vbShortTime)) >= 0 Then
                            GetMarcacionVarios = rstMarcaciones!idmarcacion
                            ostrTipoMarcacion = gstrSalida
                            Exit Do
                        End If
                    End If
                    rstMarcaciones.MoveNext
                Loop
        
            '=============================
            '*** PROCEDIMIENTO SIN HOLGURA
            '=============================
            Else
                GetMarcacionVarios = 0
                If rstMarcaciones.RecordCount = 0 Then Exit Function
                'almacena en las variables el codigo de la primera y ultima marcacion
                lngPrimeraMarcacion = rstMarcaciones!idmarcacion
                rstMarcaciones.MoveLast
                lngUltimaMarcacion = rstMarcaciones!idmarcacion
                rstMarcaciones.MoveFirst
                'Por cada compromiso de marcacion para este dia
                Do While Not rstMarcaciones.EOF
                    dblMinutosHolgura = DateDiff("n", rstMarcaciones!HoraIngresoIdeal, rstMarcaciones!HoraSalidaIdeal) / 2
                    'Si la entrada esta vacia
                    If IsNull(rstMarcaciones!HoraIngresoReal) Then
                        If lngPrimeraMarcacion = rstMarcaciones!idmarcacion Then
                            'si es la primera marcacion del dia
                            If DateDiff("n", idatHoraReal, FormatDateTime(rstMarcaciones!HoraSalidaIdeal, vbShortTime)) >= 0 Then 'llega antes de la hora de salida
                                GetMarcacionVarios = rstMarcaciones!idmarcacion
                                ostrTipoMarcacion = gstrEntrada
                                Exit Do
                            End If
                        Else
                            If DateDiff("n", idatHoraReal, FormatDateTime(DateAdd("n", -dblMinutosHolgura, rstMarcaciones!HoraIngresoIdeal), vbShortTime)) <= 0 _
                                    And DateDiff("n", idatHoraReal, FormatDateTime(DateAdd("n", dblMinutosHolgura, rstMarcaciones!HoraIngresoIdeal), vbShortTime)) >= 0 Then
                                GetMarcacionVarios = rstMarcaciones!idmarcacion
                                ostrTipoMarcacion = gstrEntrada
                                Exit Do
                            End If
                        End If
                    'Si la salida esta vacia
                    End If
                    If IsNull(rstMarcaciones!HoraSalidaReal) Then
                        If lngUltimaMarcacion = rstMarcaciones!idmarcacion Then
                            'si es la ultima marcacion del dia
                            If DateDiff("n", idatHoraReal, FormatDateTime(rstMarcaciones!HoraIngresoIdeal, vbShortTime)) <= 0 Then 'sale despues de la hora de ingreso
                                GetMarcacionVarios = rstMarcaciones!idmarcacion
                                ostrTipoMarcacion = gstrSalida
                                Exit Do
                            End If
                        Else
                            If DateDiff("n", idatHoraReal, FormatDateTime(DateAdd("n", -dblMinutosHolgura, rstMarcaciones!HoraSalidaIdeal), vbShortTime)) <= 0 _
                                        And DateDiff("n", idatHoraReal, FormatDateTime(DateAdd("n", dblMinutosHolgura, rstMarcaciones!HoraSalidaIdeal), vbShortTime)) >= 0 Then
                                GetMarcacionVarios = rstMarcaciones!idmarcacion
                                ostrTipoMarcacion = gstrSalida
                                Exit Do
                            End If
                        End If
                    End If
                    rstMarcaciones.MoveNext
                Loop
            End If
    End If
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetMarcacionVarios, mstrClassModule & " - GetMarcacionVarios Method", Err.Description, CStr(Erl))

End Function

Public Function GetJustificacion(Optional ByVal iintIdJustificacion As Long = 0, _
                            Optional ByVal iintIdTrabajador As Integer = 0, _
                            Optional ByVal iintMotivoJustificacion As Integer = 0, _
                            Optional ByVal istrRemunerada As String = "", _
                            Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                            Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                
Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
  strSQL = "SELECT DISTINCT Justificacion.IdJustificacion, Marcacion.IdTrabajador, Trabajador.Apellido+' ,'+Trabajador.Nombre AS Apellido, " _
            & " Justificacion.IdMotivoJustificacion, MotivoJustificacion.Descripcion AS MotivoJustificacion, " _
            & " Justificacion.Fecha, Remunerada, Observacion " _
            & " FROM MotivoJustificacion INNER JOIN (Justificacion INNER JOIN (Trabajador INNER JOIN Marcacion " _
            & " ON Trabajador.IdTrabajador = Marcacion.IdTrabajador) ON Justificacion.IdJustificacion = " _
            & " Marcacion.IdJustificacion) ON MotivoJustificacion.IdMotivoJustificacion = Justificacion.IdMotivoJustificacion WHERE 1=1 "
             
    If iintIdJustificacion <> 0 Then
      strSQL = strSQL & " AND Justificacion.IdJustificacion = " & iintIdJustificacion
    End If
    
    If iintIdTrabajador <> 0 Then
      strSQL = strSQL & " AND Marcacion.IdTrabajador = " & iintIdTrabajador
    End If
    
    If iintMotivoJustificacion <> 0 Then
      strSQL = strSQL & " AND Justificacion.IdMotivoJustificacion = " & iintMotivoJustificacion
    End If
    
    If istrRemunerada <> "" Then
      strSQL = strSQL & " AND Remunerada = '" & istrRemunerada & "'"
    End If
    If idatFechaInicio <> gdatFecha_Nula_RangoInicio And idatFechaFin <> gdatFecha_Nula_RangoFin Then
        strSQL = strSQL & " AND Justificacion.Fecha >= '" & idatFechaInicio & "' AND Justificacion.Fecha <= '" & idatFechaFin & "'"
    End If
    strSQL = strSQL & " ORDER BY Apellido"
    
  Set objDataAccess = CreateObject("Personal.clsData")
  Set GetJustificacion = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
  Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetJustificacion, mstrClassModule & " - GetJustificacion Method", Err.Description, CStr(Erl))

End Function

Public Function GetConsxEliminar(ByVal idatFecha As Date, _
                                Optional ByVal iarrayElimMarcacion As Variant) As Object
Dim objDataAccess As Object
Dim strSQL As String
  
On Error GoTo ErrHandler

  strSQL = "up_BuscarMarcacionxFecha '" & idatFecha & "'"
  Set objDataAccess = CreateObject("Personal.clsData")
  Set GetConsxEliminar = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
  Set objDataAccess = Nothing


Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetConsxEliminar, mstrClassModule & " - GetConsxEliminar Method", Err.Description, CStr(Erl))

End Function

Public Function NewEliminar(ByVal idatFecha As Date, _
                                ByVal iintIdMotivoJustificacion As Integer, _
                                ByVal istrRemunerada As String, _
                                ByVal iarrayMarcacion As Variant, _
                                Optional ByVal istrObservacion As String = "") As Boolean
    
Dim objDataAccess As Object
Dim intCodigo As Long
Dim intContador As Integer
Dim strSQL As String
Dim blnResult As Boolean
    
On Error GoTo ErrHandler
    intCodigo = GetSecuenciaCodigo("Justificacion")
    strSQL = "up_AgregarJustificacion '" & iintIdMotivoJustificacion & "','" & idatFecha & "','" & istrRemunerada & "','" & istrObservacion & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
    blnResult = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    If blnResult Then
        For intContador = LBound(iarrayMarcacion, 1) To UBound(iarrayMarcacion, 1)
            strSQL = "UPDATE Marcacion SET IdJustificacion = " & intCodigo & " WHERE IdMarcacion = " & iarrayMarcacion(intContador, 0)
            NewEliminar = objDataAccess.RunSQLSentence(strSQL, gStringConnection)
        Next
    End If
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    NewEliminar = False
    Call RaiseError(vbObjectError + clsMovimientos_NewEliminar, mstrClassModule & " - NewEliminar Method", Err.Description, CStr(Erl))
    
End Function

Public Function GetFaltas(Optional ByVal iintIdTrabajador As Integer = 0, _
                        Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                        Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
                                
Dim objDataAccess As Object
Dim strSQL As String
  
On Error GoTo ErrHandler

  strSQL = "up_BuscarFaltas '" & iintIdTrabajador & "','" & idatFechaInicio & "','" & idatFechaFin & "'"
  Set objDataAccess = CreateObject("Personal.clsData")
  Set GetFaltas = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
  Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetFaltas, mstrClassModule & " - GetFaltas Method", Err.Description, CStr(Erl))

End Function



Public Function GetTardanzas(Optional ByVal iintIdTrabajador As Integer = 0, _
                        Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                        Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
Dim strSQL As String
Dim objDataAccess As Object
On Error GoTo ErrHandler
    strSQL = "up_BuscarTardanzas '" & iintIdTrabajador & "','" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetTardanzas = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetTardanzas, mstrClassModule & " - GetTardanzas Method", Err.Description, CStr(Erl))
End Function

Public Function GetJornadasIncompletas(Optional ByVal iintIdTrabajador As Integer = 0, _
                        Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                        Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin) As Object
Dim strSQL As String
Dim objDataAccess As Object
On Error GoTo ErrHandler
    strSQL = "up_BuscarJornadasImcompletas '" & iintIdTrabajador & "','" & idatFechaInicio & "','" & idatFechaFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetJornadasIncompletas = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetJornadasIncompletas, mstrClassModule & " - GetJornadasIncompletas Method", Err.Description, CStr(Erl))

End Function

'Agrega registro Justificacion y modifica el idJustificacion en marcacion
Public Function AddNewJustificacion(ByVal idatFecha As Date, _
                                ByVal iintIdMotivoJustificacion As Integer, _
                                ByVal istrRemunerada As String, _
                                ByVal iarrayMarcacion As Variant, _
                                Optional ByVal istrObservacion As String = "") As Boolean
    
Dim objDataAccess As Object
Dim intCodigo As Long
Dim intContador As Integer
Dim strSQL As String
Dim blnResult As Boolean
    
On Error GoTo ErrHandler
    intCodigo = GetSecuenciaCodigo("Justificacion") + 1
    strSQL = "up_AgregarJustificacion " & intCodigo & ",'" & iintIdMotivoJustificacion & "','" & idatFecha & "','" & istrRemunerada & "','" & istrObservacion & "'"
    UpdateSecuenciaCodigo ("Justificacion")
    Set objDataAccess = CreateObject("Personal.clsData")
    blnResult = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    If blnResult Then
        For intContador = LBound(iarrayMarcacion, 1) To UBound(iarrayMarcacion, 1)
            strSQL = "UPDATE Marcacion SET IdJustificacion = " & intCodigo & " WHERE IdMarcacion = " & iarrayMarcacion(intContador, 0)
            AddNewJustificacion = objDataAccess.RunSQLSentence(strSQL, gStringConnection)
        Next
    End If
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewJustificacion = False
    Call RaiseError(vbObjectError + clsMovimientos_AddNewJustificacion, mstrClassModule & " - AddNewJustificacion Method", Err.Description, CStr(Erl))
End Function

Public Function AddNewJustificacionGrupoDia(ByVal idatFecha As Date, _
                                ByVal iintIdMotivoJustificacion As Integer, _
                                ByVal istrRemunerada As String, _
                                ByVal iarrayTrabajador As Variant, _
                                ByVal iarrayFecha As Variant, _
                                Optional ByVal istrObservacion As String = "") As Boolean
    
Dim objDataAccess As Object
Dim intCodigo As Long
Dim intContador As Integer
Dim intContadorF As Integer
Dim strSQL As String
Dim blnResult As Boolean
    
On Error GoTo ErrHandler
    intCodigo = GetSecuenciaCodigo("Justificacion") + 1
    strSQL = "up_AgregarJustificacion " & intCodigo & ",'" & iintIdMotivoJustificacion & "','" & idatFecha & "','" & istrRemunerada & "','" & istrObservacion & "'"
    UpdateSecuenciaCodigo ("Justificacion")
    Set objDataAccess = CreateObject("Personal.clsData")
    blnResult = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    If blnResult Then
        For intContador = LBound(iarrayTrabajador, 1) To UBound(iarrayTrabajador, 1)
             For intContadorF = LBound(iarrayFecha, 1) To UBound(iarrayFecha, 1)
                strSQL = "UPDATE Marcacion SET IdJustificacion = " & intCodigo & " WHERE IdTrabajador = " & iarrayTrabajador(intContador, 0) & " AND Fecha = '" & iarrayFecha(intContadorF, 0) & "' AND IdJustificacion is not Null"
                AddNewJustificacionGrupoDia = objDataAccess.RunSQLSentence(strSQL, gStringConnection)
             Next
        Next
    End If
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewJustificacionGrupoDia = False
    Call RaiseError(vbObjectError + clsMovimientos_AddNewJustificacionGrupoDia, mstrClassModule & " - AddNewJustificacionGrupoDia Method", Err.Description, CStr(Erl))
End Function

'Elimina registro Justificacion y modifica el campo relacionado (FK) en tabla Marcacion
Public Function DeleteJustificacion(ByVal iintIdJustificacion As Long, _
                        Optional ByRef ostrError As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
Dim blnResult As Boolean
    
On Error GoTo ErrHandler
    If ValidateBussines(EnumActionsMovimientos.DeleteJustificacion, ostrError, _
                        iintIdJustificacion) Then
        strSQL = "up_EliminarJustificacion '" & iintIdJustificacion & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        DeleteJustificacion = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Set objDataAccess = Nothing
    Else
        GoTo ExitMethod
    End If
Exit Function
ExitMethod:
    Set objDataAccess = Nothing
    DeleteJustificacion = False

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    DeleteJustificacion = False
    Call RaiseError(vbObjectError + clsMovimientos_DeleteJustificacion, mstrClassModule & " - DeleteJustificacion Method", Err.Description, CStr(Erl))
End Function

Public Function ExisteMesyAñoConsolidado(ByVal istrMes As String, _
                             ByVal istrAño As String) As Boolean
Dim objDataAccess As Object
Dim rstDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler
    strSQL = "up_BuscarMesyAñoConsolidado '" & istrMes & "','" & istrAño & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    If Not rstDataAccess.EOF Then GoTo ExitMethod
    ExisteMesyAñoConsolidado = False
    Set rstDataAccess = Nothing
  
Exit Function
ExitMethod:
    Set objDataAccess = Nothing
    Set rstDataAccess = Nothing
    ExisteMesyAñoConsolidado = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set rstDataAccess = Nothing
    ExisteMesyAñoConsolidado = True
    Call RaiseError(vbObjectError + clsMovimientos_ExisteMesYAño, mstrClassModule & " - DeleteJustificacion Method", Err.Description, CStr(Erl))
End Function

Public Function GetMarcacionTrabajador(Optional ByVal iintIdTrabajador As Integer = 0, _
                                       Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                                       Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin, _
                                       Optional ByVal istrApellido As String = "") As Object
Dim strSQL As String
Dim objDataAccess As Object
On Error GoTo ErrHandler
    strSQL = "SELECT IdMarcacion, Marcacion.IdTrabajador, Apellido, Nombre, Marcacion.Fecha, Marcacion.HoraIngresoIdeal, " & _
            " Marcacion.HoraSalidaIdeal, Marcacion.HoraIngresoReal, Marcacion.HoraSalidaReal " & _
            " FROM Marcacion INNER JOIN Trabajador ON Trabajador.IdTrabajador = Marcacion.IdTrabajador " & _
            " WHERE 1=1"
    If iintIdTrabajador <> 0 Then
        strSQL = strSQL & " AND Marcacion.IdTrabajador=" & iintIdTrabajador
    End If
    If istrApellido <> "" Then
        strSQL = strSQL & " AND Trabajador.Apellido LIKE '" & istrApellido & "%'"
    End If
    If idatFechaInicio <> gdatFecha_Nula_RangoInicio Then
        strSQL = strSQL & " AND Marcacion.Fecha>='" & idatFechaInicio
    End If
    If idatFechaFin <> gdatFecha_Nula_RangoFin Then
        strSQL = strSQL & "' AND Marcacion.Fecha<='" & idatFechaFin & "'"
    End If
    
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetMarcacionTrabajador = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set GetMarcacionTrabajador = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetMarcacionTrabajador, mstrClassModule & " - GetMarcacionTrabajador Method", Err.Description, CStr(Erl))
End Function

Public Function UpdateHoraEntradaYSalidaMarcacion(ByVal iintIdMarcacion As Long, _
                                                  ByVal idatHoraEntradaReal As Date, _
                                                  ByVal idatHoraSalidaReal As Date) As Boolean

Dim objDataAccess As Object
Dim strSQL As String
Dim blnResult As Boolean

On Error GoTo ErrHandler
    strSQL = "up_ModificarMarcacion '" & iintIdMarcacion & "','" & "" & "','" & "" & "','" & idatHoraEntradaReal & "','" & idatHoraSalidaReal & "','" & "" & "'"
     Set objDataAccess = CreateObject("Personal.clsData")
    If Not objDataAccess.RunStoredProcedure(strSQL, gStringConnection) Then
        UpdateHoraEntradaYSalidaMarcacion = False
        Exit Function
    End If
    UpdateHoraEntradaYSalidaMarcacion = True
    Set objDataAccess = Nothing
    
Exit Function
ErrHandler:
    UpdateHoraEntradaYSalidaMarcacion = False
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_UpdateHoraEntradaYSalidaMarcacion, mstrClassModule & " - UpdateHoraEntradaYSalidaMarcacion Method", Err.Description, CStr(Erl))
End Function

Public Function UpdateEstadoSuspendidoToNormal() As Boolean
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler
    '***Revisar Cuando sea el caso de Trabajar con Horarios Corridos
    strSQL = "up_ModificarEstadoSuspendidoToNormal '" & gstrEstadoTrabajadorNormal & "','" & gstrEstadoTrabajadorSuspendido & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    If Not objDataAccess.RunStoredProcedure(strSQL, gStringConnection) Then
        UpdateEstadoSuspendidoToNormal = False
        Set objDataAccess = Nothing
        Exit Function
    End If
    UpdateEstadoSuspendidoToNormal = True
    Set objDataAccess = Nothing
   
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateEstadoSuspendidoToNormal = False
    
    Call RaiseError(vbObjectError + clsMovimientos_UpdateEstadoSuspendidoToNormal, mstrClassModule & " - UpdateEstadoSuspendidoToNormal Method", Err.Description, CStr(Erl))
End Function

Public Function GetConsolidadoxTrabajador(ByVal iintIdTrabajador As Integer, _
                                          ByVal idatFechaInicio As Date, _
                                          ByVal idatFechaFin As Date) As Object
Dim rstDatosTrab As Object
Dim rstMarcacion As Object
Dim rstFeriado As Object
Dim objFeriado As Object
Dim rstMotivoJustificacion As Object
Dim intIdMotivoJustificacion As Integer
Dim rstSuspension As Object
Dim objDataAccess As Object
Dim dblFaltasInj As Double
Dim intFaltasInj As Integer
Dim dblTardInj As Double
Dim intTardInj As Integer
Dim dblJorIncInj As Double
Dim intJorIncInj As Integer
Dim dblFeriadosNR As Double
Dim intFeriadosNR As Integer
Dim dblSuspensiones As Double
Dim intSuspensiones As Integer
Dim TotalHorasNR As Double
Dim dblHorasTrab As Double
Dim dblFaltasJust As Double
Dim intFaltasJust As Integer
Dim dblTardJust As Double
Dim intTardJust As Integer
Dim dblJorIncJust As Double
Dim intJorIncJust As Integer
Dim dblTardJustNR As Double
Dim intTardJustNR As Integer
Dim dblJorIncJustNR As Double
Dim intJorIncJustNR As Integer
Dim dblFaltasJustNR As Double
Dim intFaltasJustNR As Integer
Dim dblFeriadosR As Double
Dim intFeriadosR As Integer
Dim intHolguraEntrada As Integer
Dim intHolguraSalida As Integer
Dim datUltimaFecha As Date
Dim rstTrabajador As Object

On Error GoTo ErrHandler
    Set rstDatosTrab = GetRecordsetDesconectado("Entero AS IdTrabajador, Texto AS Apellidos, Entero AS FID," & _
                                         " Entero AS FNRD, Doble AS SPD, Doble AS FIH, Doble AS FNRH," & _
                                         " Doble AS SPH, Entero AS FRD,Doble AS FRH, Doble AS TARDRD, Entero as TARDNRD," & _
                                         " Doble AS TARDRH,Doble as TARDNRH,Doble AS HORASTRAB, Entero AS FJNRD, Doble AS FJNRH," & _
                                         " Entero AS JIRD, Doble AS JIRH, Entero AS JIINJD, Doble AS JIINJH, Entero AS JIJUSTNRD," & _
                                         " Doble AS JIJUSTNRH,Entero AS FJRD, Doble AS FJRH")
    Set rstMarcacion = GetMarcacionesTrabajador(iintIdTrabajador, idatFechaInicio, idatFechaFin)
    Set objDataAccess = CreateObject("Personal.clsPersonal")
    intHolguraEntrada = objDataAccess.GetHolgura("Entrada")
    intHolguraSalida = objDataAccess.GetHolgura("Salida")
    Set objDataAccess = Nothing
    Set objFeriado = CreateObject("Personal.clsMovimientos")
    
    datUltimaFecha = gdatFecha_Nula
     
    Do While Not rstMarcacion.EOF
    
    If IsNull(rstMarcacion!IdJustificacion) Then
        Set rstFeriado = objFeriado.GetFeriado(, rstMarcacion!FECHA)
        If Not rstFeriado.EOF Then
            If rstFeriado!Remunerado Then  'If rstMotivoJustificacion!Remunerada = gstrHoraRemunerada Then
                Set rstFeriado = Nothing
                dblFeriadosR = dblFeriadosR + DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraSalidaIdeal)
                If datUltimaFecha <> rstMarcacion!FECHA Then
                    datUltimaFecha = rstMarcacion!FECHA
                    intFeriadosR = intFeriadosR + 1
                End If
            Else
                Set rstFeriado = Nothing
                dblFeriadosNR = dblFeriadosNR + DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraSalidaIdeal)
                If datUltimaFecha <> rstMarcacion!FECHA Then
                    datUltimaFecha = rstMarcacion!FECHA
                    intFeriadosNR = intFeriadosNR + 1
                End If
            End If
        Else
        Set rstFeriado = Nothing
            If (IsNull(rstMarcacion!HoraIngresoReal) Or IsNull(rstMarcacion!HoraSalidaReal)) Then
                dblFaltasInj = dblFaltasInj + DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraSalidaIdeal)
                intFaltasInj = intFaltasInj + 1
            Else
            'Si llega tarde
                If DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraIngresoReal) > intHolguraEntrada Then
                    dblTardInj = dblTardInj + DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraIngresoReal)
                    intTardInj = intTardInj + 1
                'Si sale Temprano
                    If DateDiff("n", rstMarcacion!HoraSalidaReal, rstMarcacion!HoraSalidaIdeal) > intHolguraSalida Then
                        dblJorIncInj = dblJorIncInj + DateDiff("n", rstMarcacion!HoraSalidaReal, rstMarcacion!HoraSalidaIdeal)
                        intJorIncInj = intJorIncInj + 1
                        dblHorasTrab = dblHorasTrab + DateDiff("n", rstMarcacion!HoraIngresoReal, rstMarcacion!HoraSalidaReal)
                    Else
                        dblHorasTrab = dblHorasTrab + DateDiff("n", rstMarcacion!HoraIngresoReal, rstMarcacion!HoraSalidaIdeal)
                    End If
                Else
                'Si sale Temprano
                    If DateDiff("n", rstMarcacion!HoraSalidaReal, rstMarcacion!HoraSalidaIdeal) > intHolguraSalida Then
                        dblJorIncInj = dblJorIncInj + DateDiff("n", rstMarcacion!HoraSalidaReal, rstMarcacion!HoraSalidaIdeal)
                        intJorIncInj = intJorIncInj + 1
                        dblHorasTrab = dblHorasTrab + DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraSalidaReal)
                    Else
                        dblHorasTrab = dblHorasTrab + DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraSalidaIdeal)
                    End If
                End If
            End If
        End If
    Else
    Set rstMotivoJustificacion = GetMotivoJustificacion(CInt(rstMarcacion!IdJustificacion))
    intIdMotivoJustificacion = rstMotivoJustificacion!IdMotivoJustificacion
    'Set objFeriado = CreateObject("Personal.clsMovimientos")
    Set rstFeriado = objFeriado.GetFeriado(, rstMarcacion!FECHA)
    'Set objFeriado = Nothing
    If Not rstFeriado.EOF Then
    
        If rstFeriado!Remunerado Then
            Set rstFeriado = Nothing
            dblFeriadosR = dblFeriadosR + DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraSalidaIdeal)
            If datUltimaFecha <> rstMarcacion!FECHA Then
                datUltimaFecha = rstMarcacion!FECHA
                intFeriadosR = intFeriadosR + 1
            End If
        Else
            Set rstFeriado = Nothing
            dblFeriadosNR = dblFeriadosNR + rstMarcacion!HoraSalidaIdeal - rstMarcacion!HoraIngresoIdeal
            If datUltimaFecha <> rstMarcacion!FECHA Then
                datUltimaFecha = rstMarcacion!FECHA
                intFeriadosNR = intFeriadosNR + 1
            End If
        End If
    Else
    Set rstFeriado = Nothing
    If rstMotivoJustificacion!Remunerada = gstrHoraRemunerada Then
            If IsNull(rstMarcacion!HoraIngresoReal) Or IsNull(rstMarcacion!HoraSalidaReal) Then
                dblFaltasJust = dblFaltasJust + DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraSalidaIdeal)
                intFaltasJust = intFaltasJust + 1
            Else
                If DateDiff("n", rstMarcacion!HoraIngresoReal, rstMarcacion!HoraIngresoIdeal) > intHolguraEntrada Then
                    dblTardJust = dblTardJust + DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraIngresoReal)
                    intTardJust = intTardJust + 1
                End If
                If DateDiff("n", rstMarcacion!HoraSalidaReal, rstMarcacion!HoraSalidaIdeal) > intHolguraSalida Then
                    dblJorIncJust = dblJorIncJust + DateDiff("n", rstMarcacion!HoraSalidaReal, rstMarcacion!HoraSalidaIdeal)
                    intJorIncJust = intJorIncJust + 1
                End If
            End If
        Else
            If (IsNull(rstMarcacion!HoraIngresoReal) Or IsNull(rstMarcacion!HoraSalidaReal)) Then
                dblFaltasJustNR = dblFaltasJustNR + DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraSalidaIdeal)
                intFaltasJustNR = intFaltasJustNR + 1
            Else
            'Si llega tarde
                If DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraIngresoReal) > intHolguraEntrada Then
                    dblTardJustNR = dblTardJustNR + DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraIngresoReal)
                    intTardJustNR = intTardJustNR + 1
                'Si sale Temprano
                    If DateDiff("n", rstMarcacion!HoraSalidaReal, rstMarcacion!HoraSalidaIdeal) > intHolguraSalida Then
                        dblJorIncJustNR = dblJorIncJustNR + DateDiff("n", rstMarcacion!HoraSalidaReal, rstMarcacion!HoraSalidaIdeal)
                        intJorIncJustNR = intJorIncJustNR + 1
                        dblHorasTrab = dblHorasTrab + DateDiff("n", rstMarcacion!HoraIngresoReal, rstMarcacion!HoraSalidaReal)
                    Else
                        dblHorasTrab = dblHorasTrab + DateDiff("n", rstMarcacion!HoraIngresoReal, rstMarcacion!HoraSalidaIdeal)
                    End If
                Else
                'Si sale Temprano
                    If DateDiff("n", rstMarcacion!HoraSalidaReal, rstMarcacion!HoraSalidaIdeal) > intHolguraSalida Then
                        dblJorIncJustNR = dblJorIncJustNR + DateDiff("n", rstMarcacion!HoraSalidaReal, rstMarcacion!HoraSalidaIdeal)
                        intJorIncJustNR = intJorIncJustNR + 1
                        dblHorasTrab = dblHorasTrab + DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraSalidaReal)
                    Else
                        dblHorasTrab = dblHorasTrab + DateDiff("n", rstMarcacion!HoraIngresoIdeal, rstMarcacion!HoraSalidaIdeal)
                    End If
                End If
            End If
        End If
    End If
    End If
    rstMarcacion.MoveNext
Loop
    Set objDataAccess = CreateObject("Personal.clsPersonal")
    Set rstTrabajador = objDataAccess.GetTrabajador(iintIdTrabajador)
    Set objDataAccess = Nothing
    
    
rstDatosTrab.AddNew
rstDatosTrab!IdTrabajador = iintIdTrabajador
rstDatosTrab!Apellidos = rstTrabajador!Apellido & ", " & StrConv(rstTrabajador!Nombre, vbProperCase)
Set rstTrabajador = Nothing
Set objDataAccess = CreateObject("Personal.clsConsolidado")
Set rstTrabajador = GetMinutoSuspensionTrabajador(iintIdTrabajador, idatFechaInicio, idatFechaFin)
rstDatosTrab!FID = intFaltasInj
rstDatosTrab!FNRD = intFeriadosNR
rstDatosTrab!SPD = rstTrabajador!Cantidad
rstDatosTrab!FIH = dblFaltasInj
rstDatosTrab!FNRH = dblFeriadosNR
rstDatosTrab!SPH = IIf(IsNull(rstTrabajador!Diferencia), 0, rstTrabajador!Diferencia)
rstDatosTrab!FRD = intFeriadosR
rstDatosTrab!FRH = dblFeriadosR
rstDatosTrab!TARDRD = intTardJust
rstDatosTrab!TARDNRD = intTardInj
rstDatosTrab!TARDRH = dblTardJust
rstDatosTrab!TARDNRH = dblTardInj
rstDatosTrab!HORASTRAB = dblHorasTrab
rstDatosTrab!FJNRD = intFaltasJustNR
rstDatosTrab!FJNRH = dblFaltasJustNR
rstDatosTrab!JIRD = intJorIncJust
rstDatosTrab!JIRH = dblJorIncJust
rstDatosTrab!JIINJD = intJorIncInj
rstDatosTrab!JIINJH = dblJorIncInj
rstDatosTrab!JIJUSTNRD = intJorIncJustNR
rstDatosTrab!JIJUSTNRH = dblJorIncJustNR
rstDatosTrab!FJRD = intFaltasJust
rstDatosTrab!FJRH = dblFaltasJust
rstDatosTrab.Update
Set GetConsolidadoxTrabajador = rstDatosTrab
Set rstDatosTrab = Nothing
Set rstTrabajador = Nothing
Set objFeriado = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set rstFeriado = Nothing
    Set objFeriado = Nothing
    Set rstDatosTrab = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetConsolidadosxTrabajador, mstrClassModule & " - GetConsolidadosxTrabajador Method", Err.Description, CStr(Erl))
End Function

Public Function GetRecordsetDesconectado(ByVal istrCampos As String) As Object
Dim strSQL As String
Dim objDataAccess As Object
On Error GoTo ErrHandler
    strSQL = "SELECT " & istrCampos & " FROM Campos"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRecordsetDesconectado = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection, DataAccessLockOptimistic, DataAccessOpenStatic)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set GetRecordsetDesconectado = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetRecordsetDesconectado, mstrClassModule & " - GetRecordsetDesconectado Method", Err.Description, CStr(Erl))
End Function

Public Function GetMotivoJustificacion(ByVal iintIdJustificacion As Long) As Object
Dim strSQL As String
Dim objDataAccess As Object
On Error GoTo ErrHandler
    strSQL = "up_BuscarMotivoJustificacion '" & iintIdJustificacion & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetMotivoJustificacion = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set GetMotivoJustificacion = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetMotivoJustificacion, mstrClassModule & " - GetMotivoJustificacion Method", Err.Description, CStr(Erl))
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
    Set objDataAccess = Nothing
    Set GetMarcacionesTrabajador = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetMarcacionesTrabajador, mstrClassModule & " - GetMarcacionesTrabajador Method", Err.Description, CStr(Erl))
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
    Call RaiseError(vbObjectError + clsMovimientos_GetMinutoSuspensionTrabajador, mstrClassModule & " - GetMinutoSuspensionTrabajador Method", Err.Description, CStr(Erl))
End Function

Public Function GetExisteConsolidado(ByVal idatFecha As Date) As Boolean
Dim strSQL As String
Dim objDataAccess As Object
Dim rstConsolidado As Object
On Error GoTo ErrHandler
    strSQL = "up_BuscarExisteConsolidado '" & idatFecha & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set rstConsolidado = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    If rstConsolidado.EOF Then
        GetExisteConsolidado = False
    Else
        GetExisteConsolidado = True
    End If
    Set rstConsolidado = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set rstConsolidado = Nothing
    GetExisteConsolidado = True
    Call RaiseError(vbObjectError + clsMovimientos_GetExisteConsolidado, mstrClassModule & " - GetExisteConsolidado Method", Err.Description, CStr(Erl))
End Function

Public Function AddNewFeriado(ByVal idateFechaFeriado As Date, _
                            ByVal iblnRemunerado As Byte, _
                            Optional ByRef ostrError As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    If ValidateBussines(EnumActionsMovimientos.AddNewFeriado, ostrError, _
                        , , , idateFechaFeriado) Then
        strSQL = "up_AgregarFeriado '" & idateFechaFeriado & "','" & iblnRemunerado & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        AddNewFeriado = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Set objDataAccess = Nothing
    Else
        GoTo ExitMethod
    End If
    
Exit Function
ExitMethod:
    Set objDataAccess = Nothing
    AddNewFeriado = False
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewFeriado = False
    Call RaiseError(vbObjectError + clsPersonal_AddNewFeriado, mstrClassModule & " - AddNewFeriado Method", Err.Description, CStr(Erl))

End Function

Public Function UpdateFeriado(ByVal iintIdFeriado As Integer, _
                              ByVal idateFechaFeriado As Date, _
                              ByVal iblnRemunerado As Byte) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_ModificarFeriado '" & iintIdFeriado & "','" & idateFechaFeriado & "','" & iblnRemunerado & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    UpdateFeriado = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateFeriado = False
    Call RaiseError(vbObjectError + clsPersonal_UpdateFeriado, mstrClassModule & " - UpdateFeriado Method", Err.Description, CStr(Erl))
End Function

Public Function DeleteFeriado(ByVal iintIdFeriado As Integer, _
                        Optional ByRef ostrError As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
On Error GoTo ErrHandler
    
    If ValidateBussines(EnumActionsMovimientos.DeleteFeriado, ostrError, _
                        , , iintIdFeriado) Then
        strSQL = "up_EliminarFeriado '" & iintIdFeriado & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        DeleteFeriado = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Set objDataAccess = Nothing
    Else
        GoTo ExitMethod
    End If
 
Exit Function
ExitMethod:
    Set objDataAccess = Nothing
    DeleteFeriado = False
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    DeleteFeriado = False
    Call RaiseError(vbObjectError + clsPersonal_DeleteFeriado, mstrClassModule & " - DeleteFeriado Method", Err.Description, CStr(Erl))
    
End Function

Public Function GetFeriado(Optional ByVal iintIdFeriado As Integer = 0, _
                           Optional ByVal idateFechaFeriado As Date = gdatFecha_Nula, _
                           Optional ByVal istrRemunerado As String = "") As Object
                            
    
Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
10  strSQL = "SELECT IdFeriado, FechaFeriado, Remunerado FROM Feriado WHERE 1=1 "

    If iintIdFeriado <> 0 Then
20      strSQL = strSQL & " AND IdFeriado = '" & iintIdFeriado & "'"
    End If
    If idateFechaFeriado <> gdatFecha_Nula Then
30      strSQL = strSQL & " AND FechaFeriado = '" & idateFechaFeriado & "'"
    End If
    If istrRemunerado <> "" Then
        strSQL = strSQL & " AND Remunerado = " & istrRemunerado
    End If
40  Set objDataAccess = CreateObject("Personal.clsData")
50  Set GetFeriado = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
60  Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetFeriado, mstrClassModule & " - GetFeriado Method", Err.Description, CStr(Erl))

End Function

'*************************************
'*** Encargado de mantener la integridad de los datos
Public Function ValidateBussines(ByVal iintAction As Integer, ByRef ostrMensajeError As String, _
                    Optional ByVal iintIdJustificacion As Long = 0, _
                    Optional ByVal iintIdSuspension As Long = 0, _
                    Optional ByVal iintIdFeriado As Integer = 0, _
                    Optional ByVal idatFechaFeriado As Date = gdatFecha_Nula) As Boolean

Dim objDataAccess As Object
Dim rstConsulta As Object
Dim strSQL As String
Dim datFechaInicio As Date
Dim strMes As String
Dim intAño As Integer
 
On Error GoTo ErrHandler
    Select Case iintAction
        
        '****************** Justificacion
        Case EnumActionsMovimientos.DeleteJustificacion
            'strSQL = "SELECT Marcacion.Fecha FROM Justificacion " & _
            " INNER JOIN Marcacion ON Justificacion.IdJustificacion = Marcacion.IdJustificacion " & _
            " WHERE Marcacion.CodigoConsolidado IS NOT NULL AND Marcacion.IdJustificacion = " & iintIdJustificacion
            strSQL = "SELECT Fecha FROM Justificacion WHERE IdJustificacion = " & iintIdJustificacion
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
            datFechaInicio = rstConsulta!FECHA
            intAño = Year(datFechaInicio)
            strMes = ObtenerMes(datFechaInicio)
            strSQL = "SELECT CodigoConsolidado FROM ConsolidadoAsistencia " & _
                " WHERE MES = '" & strMes & "' AND AÑO = " & intAño
            Set rstConsulta = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, el consolidado ha sido generado"
                GoTo ExitMethod
            End If
            
        '****************** Suspension ******************************
        Case EnumActionsMovimientos.DeleteSuspension
            strSQL = "SELECT Suspension.IdSuspension, Suspension.FechaInicio, Suspension.FechaFin " & _
                    " FROM Suspension WHERE Suspension.IdSuspension = " & iintIdSuspension
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
            datFechaInicio = rstConsulta!FechaInicio
            intAño = Year(datFechaInicio)
            strMes = ObtenerMes(datFechaInicio)
            strSQL = "SELECT CodigoConsolidado FROM ConsolidadoAsistencia " & _
                " WHERE MES = '" & strMes & "' AND AÑO = " & intAño
            Set rstConsulta = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, el consolidado ha sido generado"
                GoTo ExitMethod
            End If
        
         '****************** Feriado ******************
        Case EnumActionsMovimientos.DeleteFeriado
            strSQL = "SELECT FechaFeriado FROM Feriado WHERE IdFeriado = " & iintIdFeriado
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
            datFechaInicio = rstConsulta!FechaFeriado
            intAño = Year(datFechaInicio)
            strMes = ObtenerMes(datFechaInicio)
            strSQL = "SELECT CodigoConsolidado FROM ConsolidadoAsistencia " & _
                " WHERE MES = '" & strMes & "' AND AÑO = " & intAño
            Set rstConsulta = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, el consolidado ha sido generado"
                GoTo ExitMethod
            End If
            
        Case EnumActionsMovimientos.AddNewFeriado
            intAño = Year(idatFechaFeriado)
            strMes = ObtenerMes(idatFechaFeriado)
            strSQL = "SELECT CodigoConsolidado FROM ConsolidadoAsistencia " & _
                " WHERE MES = '" & strMes & "' AND AÑO = " & intAño
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible crear el registro, el consolidado ha sido generado"
                GoTo ExitMethod
            End If
        
    End Select
            
    Set objDataAccess = Nothing
    Set rstConsulta = Nothing
    ValidateBussines = True
    
Exit Function
ExitMethod:
    Set objDataAccess = Nothing
    Set rstConsulta = Nothing
    ValidateBussines = False

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set rstConsulta = Nothing
    ValidateBussines = False

    Err.Raise Number:=vbObjectError + clsMovimientos_ValidateBussines, Source:=mstrClassModule & " - ValidateBussines Method", Description:=Err.Description
End Function

Public Function Transaccion_Marcacion(ByVal iarrayTrabajador As Variant, ByVal FECHA As Date) As Object
Dim strSQL As String
Dim objDataAccess As Object
Dim rstDataAccess As Object

Dim rstDetalle As Object
Dim objPersonal As Object

Dim i As Integer
Dim intUboundTrabajador As Integer
On Error GoTo ErrHandler
'        Set rstDetalle = Nothing
'        Set objPersonal = CreateObject("Personal.clsGeneral")
'        Set rstDetalle = objPersonal.GetCampos("Entero As Idtrabajador, Texto As NombreCompleto, Boolean as Eliminar")
'        set objPersonal = nothing
            
    intUboundTrabajador = UBound(iarrayTrabajador)
    For i = LBound(iarrayTrabajador) To intUboundTrabajador - 1
        strSQL = "UP_TRANSACCION_MARCACION " & iarrayTrabajador(i) & ",'" & FECHA & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        objDataAccess.RunStoredProcedure strSQL, gStringConnection
    Set objDataAccess = Nothing
    Next
 
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsHorario_Transaccion_Marcacion, Source:=mstrClassModule & " - Transaccion_Marcacion Method", Description:=Err.Description
End Function

Public Function GetConsHorarioxEliminar(ByVal idatFecha As Date) As Object
Dim objDataAccess As Object
Dim strSQL As String
  
On Error GoTo ErrHandler

  strSQL = "up_BuscarMarcacionHorarioxFecha '" & idatFecha & "'"
  Set objDataAccess = CreateObject("Personal.clsData")
  Set GetConsHorarioxEliminar = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
  Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetConsHorarioxEliminar, mstrClassModule & " - GetConsHorarioxEliminar Method", Err.Description, CStr(Erl))

End Function


Public Function GetMarcacionxModificar(ByVal idatFecha As Date) As Object
Dim objDataAccess As Object
Dim strSQL As String
  
On Error GoTo ErrHandler

  strSQL = "up_BuscarMarcacionModificarxFecha '" & idatFecha & "'"
  Set objDataAccess = CreateObject("Personal.clsData")
  Set GetMarcacionxModificar = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
  Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set GetMarcacionxModificar = Nothing
    Call RaiseError(vbObjectError + clsMovimientos_GetConsHorarioxEliminar, mstrClassModule & " - GetConsHorarioxEliminar Method", Err.Description, CStr(Erl))

End Function

'Agrega registro Justificacion y modifica el idJustificacion en marcacion
'y agrega marcacion con permiso
Public Function AddNewPermiso(ByVal idatFecha As Date, _
                                ByVal istrRemunerada As String, _
                                ByVal iarrayMarcacion As Variant, _
                                Optional ByVal istrObservacion As String = "") As Boolean
    
Dim objDataAccess As Object
Dim intCodigo As Long
Dim intContador As Integer
Dim strSQL As String
Dim blnResult As Boolean
On Error GoTo ErrHandler
    intCodigo = GetSecuenciaCodigo("Justificacion") + 1
    strSQL = "up_AgregarJustificacion " & intCodigo & ",'" & gintPermiso & "','" & idatFecha & "','" & istrRemunerada & "','" & istrObservacion & "'"
    UpdateSecuenciaCodigo ("Justificacion")
    Set objDataAccess = CreateObject("Personal.clsData")
    blnResult = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    If blnResult Then
        For intContador = LBound(iarrayMarcacion, 1) To UBound(iarrayMarcacion, 1)
            strSQL = "up_ModificarMarcacionPermiso " & intCodigo & ",'" & iarrayMarcacion(intContador, 0) & "','" & FormatDateTime(iarrayMarcacion(intContador, 2), vbShortTime) & "','" & FormatDateTime(iarrayMarcacion(intContador, 3), vbShortTime) & "'"
            AddNewPermiso = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Next
    End If
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewPermiso = False
    Call RaiseError(vbObjectError + clsMovimientos_AddNewPermiso, mstrClassModule & " - AddNewPermiso Method", Err.Description, CStr(Erl))
    
End Function

Public Function UpdateJustificacion(ByVal iintIdJustificacion As Integer, _
                                ByVal iintIdMotivoJustificacion As Integer, _
                                ByVal istrRemunarada As String, _
                                ByVal istrObservacion As String, _
                                ByVal idatFecha As String) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    strSQL = "up_ModificarJustificacion '" & iintIdJustificacion & "','" & iintIdMotivoJustificacion & "','" & istrRemunarada & "','" & istrObservacion & "','" & idatFecha & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    UpdateJustificacion = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateJustificacion = False
    Call RaiseError(vbObjectError + clsMovimientos_UpdateJustificacion, mstrClassModule & " - UpdateJustificacion Method", Err.Description, CStr(Erl))
End Function




