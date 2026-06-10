VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const mstrClassModule As String = "Personal - cHorario Class Module"

Public Function AddNewHorario(ByVal idatFecha As Date, _
                                ByVal idatHoraIngreso As Date, _
                                ByVal idatHoraSalida As Date) As Boolean
    
Dim objDataAccess As Object
Dim intCodigo As Integer
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_AgregarHorario '" & idatFecha & "','" & idatHoraIngreso & "','" & idatHoraSalida & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    AddNewHorario = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    AddNewHorario = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewHorario = False
    Call RaiseError(vbObjectError + clsHorario_AddNewHorario, mstrClassModule & " - AddNewHorario Method", Err.Description, CStr(Erl))
    
End Function

Public Function UpdateHorario(ByVal iintIdHorario As Long, _
                            ByVal idatFecha As Date, _
                            ByVal idatHoraIngreso As Date, _
                            ByVal idatHoraSalida As Date) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_ModificarHorario '" & iintIdHorario & "','" & idatFecha & "','" & idatHoraIngreso & "','" & idatHoraSalida & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    UpdateHorario = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    UpdateHorario = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateHorario = False
    Call RaiseError(vbObjectError + clsHorario_UpdateHorario, mstrClassModule & " - UpdateHorario Method", Err.Description, CStr(Erl))
End Function

Public Function DeleteHorario(ByVal iintIdHorario As Long) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_EliminarHorario '" & iintIdHorario & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    DeleteHorario = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    DeleteHorario = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    DeleteHorario = False
    Call RaiseError(vbObjectError + clsHorario_DeleteHorario, mstrClassModule & " - DeleteHorario Method", Err.Description, CStr(Erl))
    
End Function

Public Function GetHorario(Optional ByVal iintIdHorario As Long = 0) As Object

Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
10  strSQL = "up_BuscarHorario '" & iintIdHorario & "'"
20  Set objDataAccess = CreateObject("Personal.clsData")
30  Set GetHorario = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
40  Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsHorario_GetHorario, mstrClassModule & " - GetHorario Method", Err.Description, CStr(Erl))

End Function

Public Function GetHorarioxTrabajador(Optional ByVal iintIdTrabajador As Integer = 0, _
                        Optional ByVal iintIdCategoria As Integer = 0, _
                        Optional ByVal iintIdCargo As Integer = 0, _
                        Optional ByVal iintIdArea As Integer = 0, _
                        Optional ByVal iintIdTipoHora As Integer = 0, _
                        Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                        Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin, _
                        Optional ByVal idatHoraInicio As Date = gdatHora_Nula_RangoInicio, _
                        Optional ByVal idatHoraFin As Date = gdatHora_Nula_RangoFin, _
                        Optional ByVal iintIdGrupo As Integer = 0, _
                        Optional ByVal iintIdHorario As Long = 0) As Object
                                
Dim objDataAccess As Object
Dim strSQL As String
Dim strCondicion As String
Dim strHoraInicio As String
Dim strHoraFin As String
Dim intRemplazarHoraInicio As Integer
Dim intRemplazarHoraFin As Integer
   
On Error GoTo ErrHandler
 
strSQL = "SELECT HorarioxTrabajador.IdHorario, HorarioxTrabajador.IdTrabajador, Apellido, Nombre, HorarioxTrabajador.Fecha, dbo.formatdate(HorarioxTrabajador.HoraIngreso, 'hh:nn tt') as HoraIngreso, " & _
            " dbo.formatdate(HorarioxTrabajador.HoraSalida,'hh:nn tt') as HoraSalida, HorarioxTrabajador.IdTipoHora, TipoHora.Descripcion AS TipoHora, " & _
            " Trabajador.IdCategoria, Categoria.Descripcion AS Categoria, Trabajador.IdCargo, Cargo.Descripcion AS Cargo, " & _
            " Trabajador.IdArea, Area.Descripcion AS Area, Trabajador.IdGrupo, Grupo.Descripcion AS Grupo " & _
            " FROM Grupo INNER JOIN ((Categoria INNER JOIN (Area INNER JOIN (Cargo INNER JOIN Trabajador ON Cargo.IdCargo = Trabajador.IdCargo) ON Area.IdArea = Trabajador.IdArea) ON Categoria.IdCategoria = Trabajador.IdCategoria) INNER JOIN (TipoHora INNER JOIN HorarioxTrabajador ON TipoHora.IdTipoHora = HorarioxTrabajador.IdTipoHora) ON Trabajador.IdTrabajador = HorarioxTrabajador.IdTrabajador) ON Grupo.IdGrupo = Trabajador.IdGrupo" & _
            " WHERE 1=1 "
             
    If iintIdTrabajador <> 0 Then
        strSQL = strSQL & " AND HorarioxTrabajador.IdTrabajador = " & iintIdTrabajador
    End If
       
    If iintIdCategoria <> 0 Then
        strSQL = strSQL & " AND Trabajador.IdCategoria = " & iintIdCategoria
    End If
    
    If iintIdCargo <> 0 Then
        strSQL = strSQL & " AND Trabajador.IdCargo = " & iintIdCargo
    End If
    
    If iintIdArea <> 0 Then
        strSQL = strSQL & " AND Trabajador.IdArea = " & iintIdArea
    End If
    
    If iintIdGrupo <> 0 Then
        strSQL = strSQL & " AND Trabajador.IdGrupo = " & iintIdGrupo
    End If
    
    If iintIdTipoHora <> 0 Then
        strSQL = strSQL & " AND HorarioxTrabajador.IdTipoHora = " & iintIdTipoHora
    End If
       
    If idatFechaInicio <> gdatFecha_Nula_RangoInicio And idatFechaFin <> gdatFecha_Nula_RangoFin Then
        strSQL = strSQL & " AND Fecha >= '" & idatFechaInicio & "' AND Fecha <= '" & idatFechaFin & "'"
    End If
    
    If idatHoraInicio <> gdatHora_Nula_RangoInicio And idatHoraFin <> gdatHora_Nula_RangoFin Then
        strHoraInicio = FormatDateTime(idatHoraInicio, vbShortTime)
        strHoraFin = FormatDateTime(idatHoraFin, vbShortTime)
        strSQL = strSQL & " AND HoraIngreso >= '" & strHoraInicio & "' AND HoraSalida <= '" & strHoraFin & "'"
    End If
    
    If iintIdHorario <> 0 Then
        strSQL = strSQL & " AND HorarioxTrabajador.IdHorario = " & iintIdHorario
    End If
    
    strSQL = strSQL & " ORDER BY HorarioxTrabajador.Fecha, Apellido "
    
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetHorarioxTrabajador = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsHorario_GetHorarioxTrabajador, mstrClassModule & " - GetHorarioxTrabajador Method", Err.Description, CStr(Erl))

End Function

Public Function UpdateMarcacion(ByVal iintIdMarcacion As Long, _
                        Optional ByVal idatHoraIngresoIdeal As Date = gdatFecha_Nula_RangoInicio, _
                        Optional ByVal idatHoraIngresoReal As Date = gdatHora_Nula_RangoInicio, _
                        Optional ByVal idatHoraSalidaIdeal As Date = gdatFecha_Nula_RangoFin, _
                        Optional ByVal idatHoraSalidaReal As Date = gdatHora_Nula_RangoFin, _
                        Optional ByVal iintTipoMarcacion As Integer) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_ModificarMarcacionManual '" & iintIdMarcacion & "','" & Format(idatHoraIngresoIdeal, "hh:mm:ss") & "','" & Format(idatHoraSalidaIdeal, "hh:mm:ss") & "','" & Format(idatHoraIngresoReal, "hh:mm:ss") & "','" & Format(idatHoraSalidaReal, "hh:mm:ss") & "'," & iintTipoMarcacion
    Set objDataAccess = CreateObject("Personal.clsData")
    UpdateMarcacion = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    UpdateMarcacion = True

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateMarcacion = False
    Call RaiseError(vbObjectError + clsHorario_UpdateMarcacion, mstrClassModule & " - UpdateMarcacion Method", Err.Description, CStr(Erl))
    
End Function

Public Function UpdateMarcacionGrupoManual(ByVal iarrayMarcacion As Variant, _
                        Optional ByVal idatHoraIngresoIdeal As Date = gdatFecha_Nula_RangoInicio, _
                        Optional ByVal idatHoraIngresoReal As Date = gdatHora_Nula_RangoInicio, _
                        Optional ByVal idatHoraSalidaIdeal As Date = gdatFecha_Nula_RangoFin, _
                        Optional ByVal idatHoraSalidaReal As Date = gdatHora_Nula_RangoFin, _
                        Optional ByVal iintTipoMarcacion As Integer) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
Dim intContador As Integer
    
On Error GoTo ErrHandler
    
    Set objDataAccess = CreateObject("Personal.clsData")
    For intContador = LBound(iarrayMarcacion, 1) To UBound(iarrayMarcacion, 1)
        strSQL = "up_ModificarMarcacionManual '" & iarrayMarcacion(intContador) & "','" & Format(idatHoraIngresoIdeal, "hh:mm:ss") & "','" & Format(idatHoraSalidaIdeal, "hh:mm:ss") & "','" & Format(idatHoraIngresoReal, "hh:mm:ss") & "','" & Format(idatHoraSalidaReal, "hh:mm:ss") & "'," & iintTipoMarcacion
        UpdateMarcacionGrupoManual = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Next
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateMarcacionGrupoManual = False
    Call RaiseError(vbObjectError + clsHorario_UpdateMarcacion, mstrClassModule & " - UpdateMarcacion Method", Err.Description, CStr(Erl))
    
End Function

Public Function UpdateMarcacionGrupoIgualIdeal(ByVal iarrayMarcacion As Variant, _
                        Optional ByVal iintTipoMarcacion As Integer) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
Dim intContador As Integer
    
On Error GoTo ErrHandler
    
    Set objDataAccess = CreateObject("Personal.clsData")
    For intContador = LBound(iarrayMarcacion, 1) To UBound(iarrayMarcacion, 1)
        strSQL = "up_ModificarMarcacionIgualIdeal '" & iarrayMarcacion(intContador) & "'," & iintTipoMarcacion
        UpdateMarcacionGrupoIgualIdeal = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Next
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateMarcacionGrupoIgualIdeal = False
    Call RaiseError(vbObjectError + clsHorario_UpdateMarcacion, mstrClassModule & " - UpdateMarcacion Method", Err.Description, CStr(Erl))
    
End Function


Public Function DeleteHorarioxTrabajador(ByVal iintIdHorario As Long, _
                                    Optional ByRef ostrError As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
If ValidateBussines(EnumActionsHorario.DeleteHorarioxTrabajador, ostrError, _
                    iintIdHorario) Then
    strSQL = "up_EliminarHorarioxTrabajador " & iintIdHorario  '& "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    DeleteHorarioxTrabajador = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    DeleteHorarioxTrabajador = True
Else
    GoTo ExitMethod
End If

Exit Function
ExitMethod:
    Set objDataAccess = Nothing
    DeleteHorarioxTrabajador = False

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    DeleteHorarioxTrabajador = False
    Call RaiseError(vbObjectError + clsHorario_DeleteHorarioxTrabajador, mstrClassModule & " - DeleteHorarioxTrabajador Method", Err.Description, CStr(Erl))
    
End Function

Public Function AddNewHorarioxTrabajador(ByVal ArrayPrimera As Variant, _
                                        ByVal ArrayTrabajador As Variant) As Boolean
    
Dim objDataAccess As Object
Dim ArraySQL() As String
Dim intContador As Integer
Dim intCodigo As Long
Dim blnResult As Boolean
On Error GoTo ErrHandler

    '#If TRANSACTION_MODE = 1 Then
    '    Dim ObjContexto As ObjectContext
    '    Set ObjContexto = GetObjectContext
    '#End If
    
    ReDim ArraySQL(LBound(ArrayPrimera) To UBound(ArrayPrimera))
    
    For intContador = LBound(ArrayPrimera, 1) To UBound(ArrayPrimera, 1) - 1
        ArraySQL(intContador) = ArrayPrimera(intContador)
        intCodigo = intCodigo + 1
    Next
    
    Set objDataAccess = CreateObject("Personal.clsData")
    blnResult = objDataAccess.RunStoredProcedureSoloDetalle(ArraySQL, gStringConnection)
    If blnResult Then
        AddNewHorarioxTrabajador = objDataAccess.RunStoredProcedureSoloDetalle(ArrayTrabajador, gStringConnection)
    Else
        GoTo ExitMethod
    End If
    Set objDataAccess = Nothing
    '#If TRANSACTION_MODE = 1 Then
    '    ObjContexto.SetComplete
    '#End If
Exit Function

ExitMethod:
    Set objDataAccess = Nothing
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    AddNewHorarioxTrabajador = False
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    AddNewHorarioxTrabajador = False
    Call RaiseError(vbObjectError + clsHorario_AddNewHorarioxTrabajador, mstrClassModule & " - AddNewHorarioxTrabajador Method", Err.Description, CStr(Erl))
End Function

Public Function AddNewHorarioxTrabajadorBK(ByVal ArrayPrimera As Variant, _
                                        ByVal ArrayTrabajador As Variant) As Boolean
    
Dim objDataAccess As Object
Dim ArraySQL() As String
Dim intContador As Integer
Dim intCodigo As Long
Dim blnResult As Boolean
On Error GoTo ErrHandler

    '#If TRANSACTION_MODE = 1 Then
    '    Dim ObjContexto As ObjectContext
    '    Set ObjContexto = GetObjectContext
    '#End If
    
    ReDim ArraySQL(LBound(ArrayPrimera) To UBound(ArrayPrimera))
    
    For intContador = LBound(ArrayPrimera, 1) To UBound(ArrayPrimera, 1) - 1
        ArraySQL(intContador) = ArrayPrimera(intContador)
        intCodigo = intCodigo + 1
    Next
    
    Set objDataAccess = CreateObject("Personal.clsData")
    blnResult = objDataAccess.RunStoredProcedureSoloDetalle(ArrayTrabajador, gStringConnection)
    If blnResult Then
        AddNewHorarioxTrabajadorBK = objDataAccess.RunStoredProcedureSoloDetalle(ArraySQL, gStringConnection)
    Else
        GoTo ExitMethod
    End If
    Set objDataAccess = Nothing
    '#If TRANSACTION_MODE = 1 Then
    '    ObjContexto.SetComplete
    '#End If
Exit Function

ExitMethod:
    Set objDataAccess = Nothing
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    AddNewHorarioxTrabajadorBK = False
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    AddNewHorarioxTrabajadorBK = False
    Call RaiseError(vbObjectError + clsHorario_AddNewHorarioxTrabajador, mstrClassModule & " - AddNewHorarioxTrabajador Method", Err.Description, CStr(Erl))
End Function

Public Function GetCompromisoMarcacion(Optional ByVal iintIdTrabajador As Integer = 0, _
                        Optional ByVal iintIdCategoria As Integer = 0, _
                        Optional ByVal iintIdCargo As Integer = 0, _
                        Optional ByVal iintIdArea As Integer = 0, _
                        Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio, _
                        Optional ByVal idatFechaFin As Date = gdatFecha_Nula_RangoFin, _
                        Optional ByVal idatHoraInicio As Date = gdatHora_Nula_RangoInicio, _
                        Optional ByVal idatHoraFin As Date = gdatHora_Nula_RangoFin, _
                        Optional ByVal iintIdMarcacion As Long = 0, _
                        Optional ByVal istrApellido As String = "") As Object
                                
                                
Dim objDataAccess As Object
Dim strSQL As String
Dim strCondicion As String
Dim strHoraInicio As String
Dim strHoraFin As String
Dim intRemplazarHoraInicio As Integer
Dim intRemplazarHoraFin As Integer
   
On Error GoTo ErrHandler
 
strSQL = "SELECT Marcacion.IdMarcacion, Marcacion.IdTrabajador, Apellido, Nombre, Marcacion.Fecha, dbo.formatdate(Marcacion.HoraIngresoIdeal,'hh:nn tt') as HoraIngresoIdeal, dbo.formatdate(Marcacion.HoraIngresoReal,'hh:nn tt') as HoraIngresoReal, " & _
            " dbo.formatdate(Marcacion.HoraSalidaIdeal,'hh:nn tt') as HoraSalidaIdeal, dbo.formatdate(Marcacion.HoraSalidaReal,'hh:nn tt') as HoraSalidaReal, Trabajador.IdCategoria, Categoria.Descripcion AS Categoria, Trabajador.IdCargo, Cargo.Descripcion AS Cargo, " & _
            " Trabajador.IdArea, Area.Descripcion AS Area " & _
            " FROM Categoria INNER JOIN (Cargo INNER JOIN (Area INNER JOIN (Trabajador INNER JOIN Marcacion ON Trabajador.IdTrabajador = Marcacion.IdTrabajador) ON Area.IdArea = Trabajador.IdArea) ON Cargo.IdCargo = Trabajador.IdCargo) ON Categoria.IdCategoria = Trabajador.IdCategoria" & _
            " WHERE 1=1 "
             
    If iintIdMarcacion <> 0 Then
        strSQL = strSQL & " AND Marcacion.IdMarcacion = " & iintIdMarcacion
    End If
    
    If iintIdTrabajador <> 0 Then
        strSQL = strSQL & " AND Marcacion.IdTrabajador = " & iintIdTrabajador
    End If
    
    If istrApellido <> "" Then
        strSQL = strSQL & " AND Apellido LIKE '" & istrApellido & "%'"
    End If
    
    If iintIdCategoria <> 0 Then
        strSQL = strSQL & " AND Trabajador.IdCategoria = " & iintIdCategoria
    End If
    
    If iintIdCargo <> 0 Then
        strSQL = strSQL & " AND Trabajador.IdCargo = " & iintIdCargo
    End If
    
    If iintIdArea <> 0 Then
        strSQL = strSQL & " AND Trabajador.IdArea = " & iintIdArea
    End If
         
    If idatFechaInicio <> gdatFecha_Nula_RangoInicio And idatFechaFin <> gdatFecha_Nula_RangoFin Then
        strSQL = strSQL & " AND Fecha >= '" & idatFechaInicio & "' AND Fecha <= '" & idatFechaFin & "'"
        'strSQL = strSQL & " AND Fecha >= " & idatFechaInicio & " AND Fecha <= " & idatFechaFin
    End If
    
    If idatHoraInicio <> gdatHora_Nula_RangoInicio And idatHoraFin <> gdatHora_Nula_RangoFin Then
'        strCondicion = Mid(idatHoraInicio, 10, 4)
'        strHoraInicio = Mid(idatHoraInicio, 1, 8)
'            If strCondicion = "a.m." And Mid(idatHoraInicio, 1, 2) = 12 Then
'                strHoraInicio = Replace(Mid(idatHoraInicio, 1, 8), "12:", "00:", 1, 1, vbTextCompare)
'            ElseIf strCondicion = "p.m." And Mid(idatHoraInicio, 1, 2) < 12 Then
'                intRemplazarHoraInicio = Mid(idatHoraInicio, 1, 2) + 12
'                strHoraInicio = Replace(Mid(idatHoraFin, 1, 8), Mid(idatHoraFin, 1, 2) & ":", intRemplazarHoraInicio & ":", 1, 1, vbTextCompare)
'            End If
'        strCondicion = Mid(idatHoraFin, 10, 4)
'        strHoraFin = Mid(idatHoraFin, 1, 8)
'            If strCondicion = "p.m." And Mid(idatHoraFin, 1, 2) < 12 Then
'                intRemplazarHoraFin = Mid(idatHoraFin, 1, 2) + 12
'                strHoraFin = Replace(Mid(idatHoraFin, 1, 8), Mid(idatHoraFin, 1, 2) & ":", intRemplazarHoraFin & ":", 1, 1, vbTextCompare)
'            ElseIf strCondicion = "a.m." And Mid(idatHoraFin, 1, 2) = 12 Then
'                strHoraFin = Replace(Mid(idatHoraFin, 1, 8), "12:", "00:", 1, 1, vbTextCompare)
'            End If
        strHoraInicio = FormatDateTime(idatHoraInicio, vbShortTime)
        strHoraFin = FormatDateTime(idatHoraFin, vbShortTime)
        strSQL = strSQL & " AND HoraIngresoIdeal >= '" & strHoraInicio & "' AND HoraSalidaIdeal <= '" & strHoraFin & "'"
'''''''''       strSQL = strSQL & " AND HoraIngresoIdeal >= " & idatHoraInicio & " AND HoraSalidaIdeal <= " & idatHoraFin
    End If
    
    strSQL = strSQL & " ORDER BY Apellido, Nombre, Marcacion.Fecha "
    
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetCompromisoMarcacion = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsHorario_GetCompromisoMarcacion, mstrClassModule & " - GetCompromisoMarcacion Method", Err.Description, CStr(Erl))

End Function

Public Function GetTrabajadorSinCompromiso() As Object
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler
    strSQL = "SELECT IdTrabajador, Apellido, Nombre FROM Trabajador " & _
            " WHERE CompromisoGenerado = " & gstrCompromisoNoGenerado & _
            " AND Estado = '" & gstrEstadoTrabajadorNormal & _
            "' ORDER BY Apellido, Nombre "
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetTrabajadorSinCompromiso = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsHorario_GetTrabajadorSinCompromiso, mstrClassModule & " - GetTrabajadorSinCompromiso Method", Err.Description, CStr(Erl))
End Function

Public Function GetCompromisoHorarioxTrabajador(ByVal iintIdTrabajador As Integer) As Object
Dim strSQL As String
Dim objDataAccess As Object
On Error GoTo ErrHandler
    strSQL = "up_BuscarHorarioxTrabajador '" & 0 & "','" & 0 & "','" & iintIdTrabajador & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetCompromisoHorarioxTrabajador = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsHorario_GetCompromisoHorarioxTrabajador, mstrClassModule & " - GetCompromisoHorarioxTrabajador Method", Err.Description, CStr(Erl))
End Function

Public Function GetHorarioxTrabajadorSinCompromiso(Optional ByVal iintIdTrabajador As Integer = 0, _
                        Optional ByVal idatFechaInicio As Date = gdatFecha_Nula_RangoInicio) As Object
                                
Dim objDataAccess As Object
Dim rstHorarioSinMarcacion As Object
Dim rstHorario As Object
Dim rstMarcacion As Object
Dim strSQLHorario As String
Dim strSQLMarcacion As String
Dim datFechaFormato As Date
   
On Error GoTo ErrHandler

datFechaFormato = idatFechaInicio
    Set rstHorarioSinMarcacion = GetRecordsetDesconectado("Entero AS IdHorario, FechaMediana AS Fecha, HoraLarga AS HoraIngreso," & _
                                                         " HoraLarga AS HoraSalida")
    'Define consulta de horarios creados para el trabajador a partir de la fecha
    strSQLHorario = "SELECT IdHorario, Fecha, HoraIngreso, HoraSalida FROM HorarioxTrabajador " & _
            " WHERE IdTrabajador=" & iintIdTrabajador & " AND Fecha >= " & datFechaFormato & " ORDER BY IdHorario"
    'Define consulta de marcaciones generadas para el trabajador horarios creados para el trabajador a partir de la fecha
    strSQLMarcacion = "SELECT IdMarcacion,IdHorario FROM Marcacion WHERE IdTrabajador=" & iintIdTrabajador & _
            " AND Fecha >= " & datFechaFormato & " ORDER BY IdHorario"
    'crea los recordset para iniciar el bucle de consulta
    Set objDataAccess = CreateObject("Personal.clsData")
    Set rstHorario = objDataAccess.GetDataFromSQLSentence(strSQLHorario, gStringConnection, DataAccessLockOptimistic)
    Set rstMarcacion = objDataAccess.GetDataFromSQLSentence(strSQLMarcacion, gStringConnection)
    Set objDataAccess = Nothing
    
    If rstHorario.RecordCount <= 0 Then Exit Function
    rstHorario.MoveFirst
    Do While Not rstHorario.EOF
        rstMarcacion.Find "IdHorario= " & rstHorario!IdHorario, , , 1
        If rstMarcacion.EOF Then
            rstHorarioSinMarcacion.AddNew
            rstHorarioSinMarcacion!IdHorario = rstHorario!IdHorario
            rstHorarioSinMarcacion!FECHA = rstHorario!FECHA
            rstHorarioSinMarcacion!HoraIngreso = rstHorario!HoraIngreso
            rstHorarioSinMarcacion!HoraSalida = rstHorario!HoraSalida
            rstHorarioSinMarcacion.Update
        End If
        rstHorario.MoveNext
    Loop
    Set rstHorario = Nothing
    Set rstMarcacion = Nothing
    Set GetHorarioxTrabajadorSinCompromiso = rstHorarioSinMarcacion
    Set rstHorarioSinMarcacion = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set rstHorarioSinMarcacion = Nothing
    Set rstHorario = Nothing
    Set rstMarcacion = Nothing
    Call RaiseError(vbObjectError + clsHorario_GetHorarioxTrabajadorSinCompromiso, mstrClassModule & " - GetHorarioxTrabajadorSinCompromiso Method", Err.Description, CStr(Erl))

End Function

Public Function AddNewCompromisoMarcacionxTrabajador(ByVal ArrayPrimera As Variant, _
                                                    ByVal ArrayTrabajador As Variant) As Boolean
    
Dim objDataAccess As Object
Dim ArraySQL() As String
Dim intContador As Long
Dim intCodigo As Long
Dim blnResult As Boolean

On Error GoTo ErrHandler
    '#If TRANSACTION_MODE = 1 Then
    '    Dim ObjContexto As ObjectContext
    '    Set ObjContexto = GetObjectContext
    '#End If
    'error al redimensionar??? esto sale porq el trabajador no tiene horario o han sido eliminados
    
    ReDim ArraySQL(LBound(ArrayPrimera) To UBound(ArrayPrimera))

    For intContador = LBound(ArrayPrimera, 1) To UBound(ArrayPrimera, 1) - 1
        ArraySQL(intContador) = ArrayPrimera(intContador)
        intCodigo = intCodigo + 1
    Next
    
    Set objDataAccess = CreateObject("Personal.clsData")
    blnResult = objDataAccess.RunStoredProcedureSoloDetalle(ArraySQL, gStringConnection)
    If blnResult Then
        AddNewCompromisoMarcacionxTrabajador = objDataAccess.RunStoredProcedureSoloDetalle(ArrayTrabajador, gStringConnection)
    Else
        GoTo ExitMethod
    End If
    Set objDataAccess = Nothing
    '#If TRANSACTION_MODE = 1 Then
    '    ObjContexto.SetComplete
    '#End If
    
Exit Function
ExitMethod:
    Set objDataAccess = Nothing
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    AddNewCompromisoMarcacionxTrabajador = False
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    AddNewCompromisoMarcacionxTrabajador = False
    Call RaiseError(vbObjectError + clsHorario_AddNewCompromisoMarcacionxTrabajador, mstrClassModule & " - AddNewCompromisoMarcacionxTrabajador Method", Err.Description, CStr(Erl))
End Function

Public Function DeleteMarcacion(ByVal iintIdTrabajador As Long, _
                                Optional ByRef ostrError As String = "") As Boolean
                             
'Public Function DeleteMarcacion(ByVal iintIdMarcacion As Long, _
'                                Optional ByRef ostrError As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler

If ValidateBussines(EnumActionsHorario.DeleteMarcacion, ostrError, _
                    , iintIdTrabajador) Then
    strSQL = "up_EliminarMarcacion '" & iintIdTrabajador & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    DeleteMarcacion = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
Else
    GoTo ExitMethod
End If

Exit Function
ExitMethod:
    Set objDataAccess = Nothing
    DeleteMarcacion = False

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    DeleteMarcacion = False
    Call RaiseError(vbObjectError + clsHorario_DeleteMarcacion, mstrClassModule & " - DeleteMarcacion Method", Err.Description, CStr(Erl))
    
End Function

Public Function UpdateHorarioxTrabajador(ByVal ilngIdHorario As Long, _
                                         ByVal idatFecha As Date, _
                                         ByVal idatIngreso As Date, _
                                         ByVal idatSalida As Date, _
                                         ByRef ostrMensajeError As String) As Boolean
Dim strSQL As String
Dim objDataAccess As Object
Dim intTotalHoras As Integer

On Error GoTo ErrHandler
    If ValidateBussines(EnumActionsHorario.UpdateHorarioxTrabajador, ostrMensajeError, ilngIdHorario) Then
        intTotalHoras = DateDiff("n", idatIngreso, idatSalida)
        strSQL = "up_ModificarHorarioxTrabajador '" & ilngIdHorario & "','" & idatFecha & "','" & FormatDateTime(idatIngreso, vbShortTime) & "','" & FormatDateTime(idatSalida, vbShortTime) & "','" & intTotalHoras & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        UpdateHorarioxTrabajador = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Set objDataAccess = Nothing
    Else
        GoTo ExitMethod
    End If

Exit Function
ExitMethod:
    Set objDataAccess = Nothing
    UpdateHorarioxTrabajador = False
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsHorario_UpdateHorarioxTrabajador, mstrClassModule & " - UpdateHorarioxTrabajador Method", Err.Description, CStr(Erl))
End Function

'modificar horario por grupo Revisar!!
Public Function UpdateHorarioxGrupo(ByVal ilngIdHorario As Long, _
                                         ByVal idatFecha As Date, _
                                         ByVal idatIngreso As Date, _
                                         ByVal idatSalida As Date, _
                                         ByRef ostrMensajeError As String) As Boolean
Dim strSQL As String
Dim objDataAccess As Object
Dim intTotalHoras As Integer

On Error GoTo ErrHandler
    If ValidateBussines(EnumActionsHorario.UpdateHorarioxTrabajador, ostrMensajeError, ilngIdHorario) Then
        intTotalHoras = DateDiff("n", idatIngreso, idatSalida)
        strSQL = "up_ModificarHorarioxTrabajador '" & ilngIdHorario & "','" & idatFecha & "','" & idatIngreso & "','" & idatSalida & "','" & intTotalHoras & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        UpdateHorarioxGrupo = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Set objDataAccess = Nothing
    Else
        GoTo ExitMethod
    End If

Exit Function
ExitMethod:
    Set objDataAccess = Nothing
    UpdateHorarioxGrupo = False
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsHorario_UpdateHorarioxGrupo, mstrClassModule & " - UpdateHorarioxGrupo Method", Err.Description, CStr(Erl))
End Function

'*************************************
'*** Encargado de mantener la integridad de los datos
Public Function ValidateBussines(ByVal iintAction As Integer, ByRef ostrMensajeError As String, _
                    Optional ByVal iintIdHorario As Long = 0, _
                    Optional ByVal iintIdMarcacion As Long = 0) As Boolean

Dim objDataAccess As Object
Dim rstConsulta As Object
Dim strSQL As String
 
On Error GoTo ErrHandler
    Select Case iintAction
        
        '****************** HorarioxTrabajador ********************
        Case EnumActionsHorario.DeleteHorarioxTrabajador
            strSQL = "SELECT * FROM Marcacion WHERE IdHorario = " & iintIdHorario
            'strSQL = "up_BuscarMarcacion '" & 0 & "','" & 0 & "','" & 0 & "','" & iintIdHorario & "','" & gdatFecha_Nula_RangoFin & "'"
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, existen código relacionados en la tabla Marcacion"
                GoTo ExitMethod
            End If
            '*********** agregado!!! ******** para eliminar horario por grupo
        Case EnumActionsHorario.DeleteHorarioxGrupo
            strSQL = "SELECT * FROM Marcacion WHERE IdHorario = " & iintIdHorario
            'strSQL = "up_BuscarMarcacion '" & 0 & "','" & 0 & "','" & 0 & "','" & iintIdHorario & "','" & gdatFecha_Nula_RangoFin & "'"
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, existen código relacionados en la tabla Marcacion"
                GoTo ExitMethod
            End If
        Case EnumActionsHorario.UpdateHorarioxTrabajador
            strSQL = "SELECT * FROM Marcacion WHERE IdHorario = " & iintIdHorario
            'strSQL = "up_BuscarMarcacion '" & 0 & "','" & 0 & "','" & 0 & "','" & iintIdHorario & "','" & gdatFecha_Nula_RangoFin & "'"
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible modificar el registro, existe registro relacionado en Marcacion"
                GoTo ExitMethod
            End If
        '****************** Marcacion **********************
        Case EnumActionsHorario.DeleteMarcacion
            strSQL = "up_BuscarMarcacion '" & iintIdMarcacion & "','" & 0 & "','" & 0 & "','" & 0 & "','" & Date & "'"
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, la fecha de Marcacion es anterior"
                GoTo ExitMethod
            End If
            
            strSQL = "SELECT IdHorario FROM Marcacion WHERE IdMarcacion = " & iintIdMarcacion & _
                            " AND (NOT ISNULL (Marcacion.HoraIngresoReal) OR NOT ISNULL (Marcacion.HoraSalidaReal))"
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, existen marcaciones registradas"
                GoTo ExitMethod
            End If
        '****************** Horario
        
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

    Err.Raise Number:=vbObjectError + clsHorario_ValidateBussines, Source:=mstrClassModule & " - ValidateBussines Method", Description:=Err.Description
End Function

Public Function DeleteHorarioYMarcacion(ByVal iarrayMarcacion As Variant) As Boolean
Dim strSQL As String
Dim objDataAccess As Object
Dim ArraySQL() As String
Dim i As Integer
Dim intUboundMarcacion As Integer
On Error GoTo ErrHandler
    intUboundMarcacion = UBound(iarrayMarcacion)
    ReDim ArraySQL(0 To 2 * intUboundMarcacion + 2)
    For i = LBound(iarrayMarcacion) To intUboundMarcacion
        ArraySQL(i) = "DELETE FROM Marcacion WHERE IdHorario=" & iarrayMarcacion(i)
    Next
    For i = intUboundMarcacion + 1 To 2 * intUboundMarcacion + 1
        ArraySQL(i) = "DELETE FROM HorarioxTrabajador WHERE IdHorario=" & iarrayMarcacion(i - intUboundMarcacion - 1)
    Next
    Set objDataAccess = CreateObject("Personal.clsData")
    DeleteHorarioYMarcacion = objDataAccess.RunSQLSoloDetalle(ArraySQL, gStringConnection)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    DeleteHorarioYMarcacion = False
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsHorario_DeleteHorarioYMarcacion, Source:=mstrClassModule & " - DeleteHorarioYMarcacion Method", Description:=Err.Description
End Function

Public Function GetHorarioGrupo(ByVal idatFechaInicio As Date, _
                                ByVal idatFechaFin As Date, _
                                Optional ByVal iintIdGrupo As Integer = 0, _
                                Optional ByVal iintIdTrabajador As Integer = 0) As Object
Dim strSQL As String
Dim objDataAccess As Object
On Error GoTo ErrHandler
    strSQL = "up_BuscarHorarioxGrupo '" & idatFechaInicio & "','" & idatFechaFin & "','" & iintIdGrupo & "','" & iintIdTrabajador & "',"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetHorarioGrupo = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Set GetHorarioGrupo = Nothing
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsHorario_GetHorarioGrupo, Source:=mstrClassModule & " - GetHorarioGrupo Method", Description:=Err.Description
End Function

Public Function GetRecordsetDesconectado(ByVal istrCampos As String) As Object
Dim strSQL As String
Dim objDataAccess As Object
    strSQL = "SELECT " & istrCampos & " FROM Campos"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetRecordsetDesconectado = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection, DataAccessLockOptimistic, DataAccessOpenStatic)
    Set objDataAccess = Nothing
End Function

Public Function GetMarcacionConHorario(ByVal iintIdHorario As Long) As Object
Dim strSQL As String
Dim objDataAccess As Object
    strSQL = "up_BuscarMarcacion '" & 0 & "','" & 0 & "','" & 0 & "','" & iintIdHorario & "','" & gdatFecha_Nula_RangoFin & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetMarcacionConHorario = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
End Function


