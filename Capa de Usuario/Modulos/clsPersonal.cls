VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const mstrClassModule As String = "Personal - cPersonal Class Module"

Public Function GetTrabajador(Optional ByVal iintIdTrabajador As Integer = 0, _
                        Optional ByVal istrTarjeta As String = "", _
                        Optional ByVal istrNombre As String = "", _
                        Optional ByVal istrApellido As String = "", _
                        Optional ByVal istrDni As String = "", _
                        Optional ByVal iintIdCategoria As Integer = 0, _
                        Optional ByVal iintIdCargo As Integer = 0, _
                        Optional ByVal iintIdArea As Integer = 0, _
                        Optional ByVal istrFoto As String = "", _
                        Optional ByVal istrEstado As String = gstrEstadoTrabajadorNormal, _
                        Optional ByVal idatFechaIngresoInferior As Date = gdatFechaInicioNull, _
                        Optional ByVal idatFechaIngresoSuperior As Date = gdatFechaFinNull, _
                        Optional ByVal idatFechaNacimientoInferior As Date = gdatFechaInicioNull, _
                        Optional ByVal idatFechaNacimientoSuperior As Date = gdatFechaFinNull, _
                        Optional ByVal iintIdGrupo As Integer = 0) As Object
                                
Dim objDataAccess As clsDatos
Dim strSQL As String
   
On Error GoTo ErrHandler
 
strSQL = "SELECT IdTrabajador, Nombre, Apellido, Tarjeta, Dni, Trabajador.IdCategoria, Trabajador.IdCargo, " & _
            " Trabajador.IdArea, Categoria.Descripcion AS Categoria, Cargo.Descripcion AS Cargo, Area.Descripcion AS Area, Foto, " & _
            " FechaIngreso, FechaNacimiento, Telefono, Direccion, Estado, Trabajador.IdGrupo, Grupo.Descripcion As Grupo,(Apellido+' '+Nombre) as Nombres, Trabajador.Huella " & _
            " FROM Grupo INNER JOIN (Categoria INNER JOIN (Area INNER JOIN (Cargo INNER JOIN Trabajador " & _
            " ON Cargo.IdCargo = Trabajador.IdCargo) ON Area.IdArea = Trabajador.IdArea) ON Categoria.IdCategoria " & _
            " = Trabajador.IdCategoria) ON Grupo.IdGrupo = Trabajador.IdGrupo " & _
            " WHERE 1=1 "
             
    If iintIdTrabajador <> 0 Then
        strSQL = strSQL & " AND IdTrabajador = " & iintIdTrabajador
    End If
    
    If istrTarjeta <> "" Then
        strSQL = strSQL & " AND Tarjeta = '" & istrTarjeta & "'"
    End If
    
    If istrNombre <> "" Then
        strSQL = strSQL & " AND Nombre LIKE '" & istrNombre & "%'"
    End If
    
    If istrApellido <> "" Then
        strSQL = strSQL & " AND Apellido LIKE '" & istrApellido & "%'"
    End If
    
    If istrDni <> "" Then
        strSQL = strSQL & " AND Dni = '" & istrDni & "'"
    End If
    
    If iintIdCategoria <> 0 Then
        strSQL = strSQL & " AND Trabajador.IdCategoria = " & iintIdCategoria
    End If
    
    If iintIdCargo <> 0 Then
        strSQL = strSQL & " AND Trabajador.IdCargo = " & iintIdCargo
    End If
    
    If iintIdGrupo <> 0 Then
        strSQL = strSQL & " AND Trabajador.IdGrupo = " & iintIdGrupo
    End If
    
    If iintIdArea <> 0 Then
        strSQL = strSQL & " AND Trabajador.IdArea = " & iintIdArea
    End If
    
    If istrFoto <> "" Then
        strSQL = strSQL & " AND Foto = '" & istrFoto & "'"
    End If
    
    If istrEstado <> "" Then
        strSQL = strSQL & " AND Estado = '" & istrEstado & "'"
    End If
    
    If idatFechaIngresoInferior <> gdatFecha_Nula_RangoInicio And idatFechaIngresoSuperior <> gdatFecha_Nula_RangoFin Then
        strSQL = strSQL & " AND FechaIngreso >= '" & idatFechaIngresoInferior & "' AND FechaIngreso <= '" & idatFechaIngresoSuperior & "'"
    End If
    
    If idatFechaNacimientoInferior <> gdatFecha_Nula_RangoInicio And idatFechaNacimientoSuperior <> gdatFecha_Nula_RangoFin Then
        strSQL = strSQL & " AND FechaNacimiento >= '" & idatFechaNacimientoInferior & "' AND FechaNacimiento <= '" & idatFechaNacimientoSuperior & "'"
    End If
    
    strSQL = strSQL & " ORDER BY Apellido, Nombre"
    
    Set objDataAccess = New clsDatos
    Set GetTrabajador = objDataAccess.GetDataFromSQLSentence(strSQL, gstrConexion)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    'Call RaiseError(vbObjectError + "clsPersonal_GetTrabajador", mstrClassModule & " - GetTrabajador Method", Err.Description, CStr(Erl))

End Function

' modifica registro de marcacion, agregando el ingreso o salida real
Public Function UpdateMarcacion(ByVal iintIdMarcacion As Long, _
                                ByVal idatHoraReal As Date, _
                                ByVal istrTipoMarcacion As String) As Boolean
    
Dim objDataAccess As clsDatos
Dim strSQL As String
    
On Error GoTo ErrHandler
    strSQL = "up_ModificarMarcacion '" & iintIdMarcacion & "','" & FormatDateTime(idatHoraReal, vbShortTime) & "','" & FormatDateTime(idatHoraReal, vbShortTime) & "','" & FormatDateTime(idatHoraReal, vbShortTime) & "','" & FormatDateTime(idatHoraReal, vbShortTime) & "','" & istrTipoMarcacion & "'"
    Set objDataAccess = New clsDatos
    UpdateMarcacion = objDataAccess.RunStoredProcedure(strSQL, gstrConexion)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateMarcacion = False
    'Call RaiseError(vbObjectError + "clsMovimientos_UpdateMarcacion", mstrClassModule & " - UpdateMarcacion Method", Err.Description, CStr(Erl))
End Function


Public Function GetUnicoTrabajador(ByVal iintIdTrabajador As Integer) As Object
                                
Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
    strSQL = "up_BuscarUnicoTrabajador '" & iintIdTrabajador & "'"
    Set objDataAccess = New clsDatos
    Set GetUnicoTrabajador = objDataAccess.GetDataFromStoredProcedure(strSQL, gstrConexion)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    'Call RaiseError(vbObjectError + clsPersonal_GetUnicoTrabajador, mstrClassModule & " - GetUnicoTrabajador Method", Err.Description, CStr(Erl))

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
    
    Set objDataAccess = New clsDatos
    Set GetMarcacion = objDataAccess.GetDataFromSQLSentence(strSQL, gstrConexion)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    'Call RaiseError(vbObjectError + clsMovimientos_GetMarcacion, mstrClassModule & " - GetMarcacion Method", Err.Description, CStr(Erl))

End Function


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
            
    Set objDataAccess = New clsDatos
    Set rstMarcaciones = objDataAccess.GetDataFromSQLSentence(strSQL, gstrConexion)
    Set objDataAccess = Nothing
    '*** OBTENER MODO DE CAPTURA DE MARCACION Y DE EXISTIR EL TOTAL DE MINUTOS
    strSQL = "SELECT Tolerancia, Tiempo, TiempoPrimeraMarcacion, HoraPrimeraMarcacion from Cliente"
    Set objDataAccess = New clsDatos
    Set rstMarcacionesTolerancia = objDataAccess.GetDataFromSQLSentence(strSQL, gstrConexion)
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
    'Call RaiseError(vbObjectError + clsMovimientos_GetMarcacionVarios, mstrClassModule & " - GetMarcacionVarios Method", Err.Description, CStr(Erl))

End Function
