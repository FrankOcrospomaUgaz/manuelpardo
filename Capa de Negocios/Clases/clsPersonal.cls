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
Attribute VB_Exposed = True
Option Explicit
Private Const mstrClassModule As String = "Personal - cPersonal Class Module"

Public Function AddNewArea(ByVal istrDescripcion As String) As Boolean
    
Dim objDataAccess As Object
Dim intCodigo As Integer
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_AgregarArea '" & istrDescripcion & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    AddNewArea = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    AddNewArea = True
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewArea = False
    Call RaiseError(vbObjectError + clsPersonal_AddNewArea, mstrClassModule & " - AddNewArea Method", Err.Description, CStr(Erl))
End Function

Public Function UpdateArea(ByVal iintIdArea As Integer, _
                            ByVal istrDescripcion As String) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_ModificarArea '" & iintIdArea & "','" & istrDescripcion & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    UpdateArea = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    UpdateArea = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateArea = False
    Call RaiseError(vbObjectError + clsPersonal_UpdateArea, mstrClassModule & " - UpdateArea Method", Err.Description, CStr(Erl))
End Function

Public Function DeleteArea(ByVal iintIdArea As Integer, _
                        Optional ByRef ostrError As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    If ValidateBussines(EnumActionsPersonal.Delete_Area, ostrError, , iintIdArea) Then
        strSQL = "up_EliminarArea '" & iintIdArea & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        DeleteArea = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Set objDataAccess = Nothing
    Else
        DeleteArea = False
        Exit Function
    End If
    DeleteArea = True
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    DeleteArea = False
    Call RaiseError(vbObjectError + clsPersonal_DeleteArea, mstrClassModule & " - DeleteArea Method", Err.Description, CStr(Erl))
End Function

Public Function GetArea(Optional ByVal iintIdArea As Integer = 0) As Object
                            
Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
10  strSQL = "up_BuscarArea '" & iintIdArea & "'"
20  Set objDataAccess = CreateObject("Personal.clsData")
30  Set GetArea = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
40  Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetArea, mstrClassModule & " - GetArea Method", Err.Description, CStr(Erl))
End Function

Public Function AddNewCargo(ByVal istrDescripcion As String) As Boolean
    
Dim objDataAccess As Object
Dim intCodigo As Integer
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_AgregarCargo '" & istrDescripcion & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    AddNewCargo = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    AddNewCargo = True
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewCargo = False
    Call RaiseError(vbObjectError + clsPersonal_AddNewCargo, mstrClassModule & " - AddNewCargo Method", Err.Description, CStr(Erl))
    
End Function

Public Function UpdateCargo(ByVal iintIdCargo As Integer, _
                            ByVal istrDescripcion As String) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_ModificarCargo '" & iintIdCargo & "','" & istrDescripcion & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    UpdateCargo = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    UpdateCargo = True
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateCargo = False
    Call RaiseError(vbObjectError + clsPersonal_UpdateCargo, mstrClassModule & " - UpdateCargo Method", Err.Description, CStr(Erl))
    
End Function

Public Function DeleteCargo(ByVal iintIdCargo As Integer, _
                        Optional ByRef ostrError As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    If ValidateBussines(EnumActionsPersonal.Delete_Cargo, ostrError, , , iintIdCargo) Then
        strSQL = "up_EliminarCargo '" & iintIdCargo & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        DeleteCargo = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Set objDataAccess = Nothing
    Else
        DeleteCargo = False
        Exit Function
    End If
    DeleteCargo = True
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    DeleteCargo = False
    Call RaiseError(vbObjectError + clsPersonal_DeleteCargo, mstrClassModule & " - DeleteCargo Method", Err.Description, CStr(Erl))
End Function

Public Function GetCargo(Optional ByVal iintIdCargo As Integer = 0) As Object
                                
Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
10  strSQL = "up_BuscarCargo '" & iintIdCargo & "'"
20  Set objDataAccess = CreateObject("Personal.clsData")
30  Set GetCargo = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
40  Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetCargo, mstrClassModule & " - GetCargo Method", Err.Description, CStr(Erl))

End Function

Public Function AddNewCategoria(ByVal istrDescripcion As String) As Boolean
    
Dim objDataAccess As Object
Dim intCodigo As Integer
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_AgregarCategoria '" & istrDescripcion & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    AddNewCategoria = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    AddNewCategoria = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewCategoria = False
    Call RaiseError(vbObjectError + clsPersonal_AddNewCategoria, mstrClassModule & " - AddNewCategoria Method", Err.Description, CStr(Erl))
    
End Function

Public Function UpdateCategoria(ByVal iintIdCategoria As Integer, _
                            ByVal istrDescripcion As String) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_ModificarCategoria '" & iintIdCategoria & "','" & istrDescripcion & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    UpdateCategoria = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    UpdateCategoria = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateCategoria = False
    Call RaiseError(vbObjectError + clsPersonal_UpdateCategoria, mstrClassModule & " - UpdateCategoria Method", Err.Description, CStr(Erl))
    
End Function

Public Function DeleteCategoria(ByVal iintIdCategoria As Integer, _
                        Optional ByRef ostrError As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    If ValidateBussines(EnumActionsPersonal.Delete_Categoria, ostrError, , , , , iintIdCategoria) Then
        strSQL = "up_EliminarCategoria '" & iintIdCategoria & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        DeleteCategoria = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Set objDataAccess = Nothing
    Else
        DeleteCategoria = False
        Exit Function
    End If
    DeleteCategoria = True
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    DeleteCategoria = False
    Call RaiseError(vbObjectError + clsPersonal_DeleteCategoria, mstrClassModule & " - DeleteCategoria Method", Err.Description, CStr(Erl))
    
End Function

Public Function GetCategoria(Optional ByVal iintIdCategoria As Integer = 0) As Object
    
Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
10  strSQL = "up_BuscarCategoria '" & iintIdCategoria & "'"
20  Set objDataAccess = CreateObject("Personal.clsData")
30  Set GetCategoria = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
40  Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetCategoria, mstrClassModule & " - GetCategoria Method", Err.Description, CStr(Erl))

End Function

Public Function AddNewTipoHora(ByVal istrDescripcion As String, _
                            ByVal istrRemunerada As String) As Boolean
    
Dim objDataAccess As Object
Dim intCodigo As Integer
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_AgregarTipoHora '" & istrDescripcion & "','" & istrRemunerada & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    AddNewTipoHora = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    AddNewTipoHora = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewTipoHora = False
    Call RaiseError(vbObjectError + clsPersonal_AddNewTipoHora, mstrClassModule & " - AddNewTipoHora Method", Err.Description, CStr(Erl))
    
End Function

Public Function UpdateTipoHora(ByVal iintIdTipoHora As Integer, _
                            ByVal istrDescripcion As String, _
                            ByVal istrRemunerada As String) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_ModificarTipoHora '" & iintIdTipoHora & "','" & istrDescripcion & "','" & istrRemunerada & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    UpdateTipoHora = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    UpdateTipoHora = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateTipoHora = False
    Call RaiseError(vbObjectError + clsPersonal_UpdateTipoHora, mstrClassModule & " - UpdateTipoHora Method", Err.Description, CStr(Erl))
    
End Function

Public Function DeleteTipoHora(ByVal iintIdTipoHora As Integer, _
                        Optional ByRef ostrError As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    If ValidateBussines(EnumActionsPersonal.Delete_TipoHora, ostrError, , , , , , iintIdTipoHora) Then
        strSQL = "up_EliminarTipoHora '" & iintIdTipoHora & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        DeleteTipoHora = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Set objDataAccess = Nothing
    Else
        DeleteTipoHora = False
        Exit Function
    End If
    DeleteTipoHora = True
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    DeleteTipoHora = False
    Call RaiseError(vbObjectError + clsPersonal_DeleteTipoHora, mstrClassModule & " - DeleteTipoHora Method", Err.Description, CStr(Erl))
    
End Function

Public Function GetTipoHora(Optional ByVal iintIdTipoHora As Integer = 0, _
                       Optional ByVal istrDescripcion As String = "", _
                       Optional ByVal istrRemunerada As String = "") As Object
                                
Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
10  strSQL = "SELECT IdTipoHora, Descripcion, Remunerada FROM TipoHora WHERE 1=1 "
             
    If iintIdTipoHora <> 0 Then
20      strSQL = strSQL & " AND IdTipoHora = " & iintIdTipoHora
    End If
    If istrDescripcion <> "" Then
30      strSQL = strSQL & " AND Descripcion = '" & istrDescripcion & "'"
    End If
    
    If istrRemunerada <> "" Then
35      strSQL = strSQL & " AND Remunerada = '" & istrRemunerada & "'"
    End If
    
40  Set objDataAccess = CreateObject("Personal.clsData")
50  Set GetTipoHora = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
60  Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetTipoHora, mstrClassModule & " - GetTipoHora Method", Err.Description, CStr(Erl))

End Function

Public Function AddNewMotivoJustificacion(ByVal istrDescripcion As String) As Boolean
    
Dim objDataAccess As Object
Dim intCodigo As Integer
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_AgregarMotivoJustificacion '" & istrDescripcion & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    AddNewMotivoJustificacion = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    AddNewMotivoJustificacion = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewMotivoJustificacion = False
    Call RaiseError(vbObjectError + clsPersonal_AddNewMotivoJustificacion, mstrClassModule & " - AddNewMotivoJustificacion Method", Err.Description, CStr(Erl))
    
End Function

Public Function UpdateMotivoJustificacion(ByVal iintIdMotivoJustificacion As Integer, _
                            ByVal istrDescripcion As String) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_ModificarMotivoJustificacion '" & iintIdMotivoJustificacion & "','" & istrDescripcion & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    UpdateMotivoJustificacion = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    UpdateMotivoJustificacion = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateMotivoJustificacion = False
    Call RaiseError(vbObjectError + clsPersonal_UpdateMotivoJustificacion, mstrClassModule & " - UpdateMotivoJustificacion Method", Err.Description, CStr(Erl))
    
End Function

Public Function DeleteMotivoJustificacion(ByVal iintIdMotivoJustificacion As Integer, _
                        Optional ByRef ostrError As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    If ValidateBussines(EnumActionsPersonal.Delete_MotivoJustificacion, ostrError, , , , , , , iintIdMotivoJustificacion) Then
        strSQL = "up_EliminarMotivoJustificacion '" & iintIdMotivoJustificacion & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        DeleteMotivoJustificacion = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Set objDataAccess = Nothing
    Else
        DeleteMotivoJustificacion = False
        Exit Function
    End If
    DeleteMotivoJustificacion = True
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    DeleteMotivoJustificacion = False
    Call RaiseError(vbObjectError + clsPersonal_DeleteMotivoJustificacion, mstrClassModule & " - DeleteMotivoJustificacion Method", Err.Description, CStr(Erl))
    
End Function

Public Function GetMotivoJustificacion(Optional ByVal iintIdMotivoJustificacion As Integer = 0) As Object

Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
10  strSQL = "up_BuscarMotivo '" & iintIdMotivoJustificacion & "'"
20  Set objDataAccess = CreateObject("Personal.clsData")
30  Set GetMotivoJustificacion = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
40  Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetMotivoJustificacion, mstrClassModule & " - GetMotivoJustificacion Method", Err.Description, CStr(Erl))

End Function

Public Function AddNewGrupo(ByVal istrDescripcion As String) As Boolean
    
Dim objDataAccess As Object
Dim intCodigo As Integer
Dim strSQL As String
    
On Error GoTo ErrHandler

    strSQL = "up_AgregarGrupo '" & istrDescripcion & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    AddNewGrupo = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    AddNewGrupo = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewGrupo = False
    Call RaiseError(vbObjectError + clsPersonal_AddNewGrupo, mstrClassModule & " - AddNewGrupo Method", Err.Description, CStr(Erl))
    
End Function

Public Function UpdateGrupo(ByVal iintIdGrupo As Integer, _
                            ByVal istrDescripcion As String) As Boolean
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_ModificarGrupo '" & iintIdGrupo & "','" & istrDescripcion & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    UpdateGrupo = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    UpdateGrupo = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateGrupo = False
    Call RaiseError(vbObjectError + clsPersonal_UpdateGrupo, mstrClassModule & " - UpdateGrupo Method", Err.Description, CStr(Erl))
    
End Function

Public Function DeleteGrupo(ByVal iintIdGrupo As Integer, _
                        Optional ByRef ostrError As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    If ValidateBussines(EnumActionsPersonal.Delete_Grupo, ostrError, , , , iintIdGrupo) Then
        strSQL = "up_EliminarGrupo '" & iintIdGrupo & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        DeleteGrupo = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Set objDataAccess = Nothing
    Else
        DeleteGrupo = False
        Exit Function
    End If
    DeleteGrupo = True
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    DeleteGrupo = False
    Call RaiseError(vbObjectError + clsPersonal_DeleteGrupo, mstrClassModule & " - DeleteGrupo Method", Err.Description, CStr(Erl))
    
End Function

Public Function GetGrupo(Optional ByVal iintIdGrupo As Integer = 0) As Object

Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
10  strSQL = "up_BuscarGrupo '" & iintIdGrupo & "'"
20  Set objDataAccess = CreateObject("Personal.clsData")
30  Set GetGrupo = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
40  Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetGrupo, mstrClassModule & " - GetGrupo Method", Err.Description, CStr(Erl))

End Function

Public Function AddNewTrabajador(ByVal istrTarjeta As String, _
                            ByVal istrNombre As String, _
                            ByVal istrApellido As String, _
                            ByVal istrDni As String, _
                            ByVal iintIdCategoria As Integer, _
                            ByVal iintIdCargo As Integer, _
                            ByVal iintIdArea As Integer, _
                            Optional ByVal iintIdGrupo As Integer = gintGrupoxDefecto, _
                            Optional ByVal istrFoto As String = "", _
                            Optional ByVal idatFechaIngreso As Date = gdatFecha_Nula, _
                            Optional ByVal idatFechaNacimiento As Date = gdatFecha_Nula, _
                            Optional ByVal istrTelefono As String = "", _
                            Optional ByVal istrDireccion As String = "", _
                            Optional ByVal istrHuella As String = "", _
                            Optional ByVal istrIndTarjeta As String = "") As Boolean
Dim objDataAccess As Object
Dim intCodigo As Integer
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_AgregarTrabajador '" & iintIdGrupo & "','" & iintIdCategoria & "','" & iintIdArea & _
             "','" & iintIdCargo & "','" & istrTarjeta & "','" & istrApellido & "','" & istrNombre & "','" & istrDni & _
             "','" & istrTelefono & "','" & istrDireccion & "','" & idatFechaNacimiento & "','" & idatFechaIngreso & _
             "','" & istrFoto & "','" & istrHuella & "','" & istrIndTarjeta & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    AddNewTrabajador = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    AddNewTrabajador = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewTrabajador = False
    Call RaiseError(vbObjectError + clsPersonal_AddNewTrabajador, mstrClassModule & " - AddNewTrabajador Method", Err.Description, CStr(Erl))
    
End Function

Public Function UpdateTrabajador(ByVal iintIdTrabajador As Integer, _
                            ByVal istrTarjeta As String, _
                            ByVal istrNombre As String, _
                            ByVal istrApellido As String, _
                            ByVal istrDni As String, _
                            ByVal iintIdCategoria As Integer, _
                            ByVal iintIdCargo As Integer, _
                            ByVal iintIdArea As Integer, _
                            ByVal istrEstado As String, _
                            Optional ByVal iintIdGrupo As Integer = gintGrupoxDefecto, _
                            Optional ByVal istrFoto As String = "", _
                            Optional ByVal idatFechaIngreso As Date = gdatFecha_Nula, _
                            Optional ByVal idatFechaNacimiento As Date = gdatFecha_Nula, _
                            Optional ByVal istrTelefono As String = "", _
                            Optional ByVal istrDireccion As String = "", _
                            Optional ByVal istrHuella As String = "", _
                            Optional ByVal istrIndTarjeta As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_ModificarTrabajador '" & iintIdTrabajador & "','" & iintIdGrupo & "','" & iintIdCategoria & _
                "','" & iintIdArea & "','" & iintIdCargo & "','" & istrTarjeta & "','" & istrApellido & _
                "','" & istrNombre & "','" & istrDni & "','" & istrTelefono & "','" & istrDireccion & _
                "','" & idatFechaNacimiento & "','" & idatFechaIngreso & "','" & istrFoto & "','" & istrEstado & "','" & Replace(istrHuella, "'", "''") & "','" & istrIndTarjeta & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    UpdateTrabajador = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    UpdateTrabajador = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateTrabajador = False
    Call RaiseError(vbObjectError + clsPersonal_UpdateTrabajador, mstrClassModule & " - UpdateTrabajador Method", Err.Description, CStr(Erl))
    
End Function

Public Function DeleteTrabajador(ByVal iintIdTrabajador As Integer, _
                        Optional ByRef ostrError As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    If ValidateBussines(EnumActionsPersonal.Delete_Trabajador, ostrError, iintIdTrabajador) Then
        strSQL = "up_EliminarTrabajador '" & iintIdTrabajador & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        DeleteTrabajador = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Set objDataAccess = Nothing
    Else
        DeleteTrabajador = False
        Exit Function
    End If
    DeleteTrabajador = True
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    DeleteTrabajador = False
    Call RaiseError(vbObjectError + clsPersonal_DeleteTrabajador, mstrClassModule & " - DeleteTrabajador Method", Err.Description, CStr(Erl))
    
End Function

Public Function GetUnicoTrabajador(ByVal iintIdTrabajador As Integer) As Object
                                
Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
    strSQL = "up_BuscarUnicoTrabajador '" & iintIdTrabajador & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetUnicoTrabajador = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetUnicoTrabajador, mstrClassModule & " - GetUnicoTrabajador Method", Err.Description, CStr(Erl))

End Function

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
                        Optional ByVal idatFechaIngresoInferior As Date = gdatFecha_Nula_RangoInicio, _
                        Optional ByVal idatFechaIngresoSuperior As Date = gdatFecha_Nula_RangoFin, _
                        Optional ByVal idatFechaNacimientoInferior As Date = gdatFecha_Nula_RangoInicio, _
                        Optional ByVal idatFechaNacimientoSuperior As Date = gdatFecha_Nula_RangoFin, _
                        Optional ByVal iintIdGrupo As Integer = 0) As Object
                                
Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
strSQL = "SELECT IdTrabajador, Nombre, Apellido, Tarjeta, Dni, Trabajador.IdCategoria, Trabajador.IdCargo, " & _
            " Trabajador.IdArea, Categoria.Descripcion AS Categoria, Cargo.Descripcion AS Cargo, Area.Descripcion AS Area, Foto, " & _
            " FechaIngreso, FechaNacimiento, Telefono, Direccion, Estado, Trabajador.IdGrupo, Grupo.Descripcion As Grupo,(Apellido+' '+Nombre) as Nombres, ISNULL(Trabajador.Huella,'') as Huella, ISNULL(IndTarjeta,'') as IndTarjeta, horario " & _
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
    
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetTrabajador = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetTrabajador, mstrClassModule & " - GetTrabajador Method", Err.Description, CStr(Erl))

End Function

Public Function GetConsolidadoAsistencia(Optional ByVal istrMes As String = "", _
                                         Optional ByVal istrAño As String = "") As Object

Dim objDataAccess As Object
Dim strSQL As String
On Error GoTo ErrHandler
    strSQL = "SELECT * FROM ConsolidadoAsistencia WHERE 1=1"
    If istrMes <> "" Then
        strSQL = strSQL = " and Mes='" & istrMes & "'"
    End If
    If istrAño <> "" Then
        strSQL = strSQL = " and Año='" & istrAño & "'"
    End If
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetConsolidadoAsistencia = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetConsolidadoAsistencia, mstrClassModule & " - GetConsolidadoAsistencia Method", Err.Description, CStr(Erl))
End Function

Public Function UpdateHolgura(ByVal iintEntrada As Integer, ByVal iintSalida As Integer) As Boolean
Dim objDataAccess As Object
Dim strSQL As String
On Error GoTo ErrHandler

'#If TRANSACTION_MODE = 1 Then
'    Dim ObjContexto As ObjectContext
'    Set ObjContexto = GetObjectContext
'#End If

    strSQL = "up_ModificarHolgura '" & iintEntrada & "','" & "Entrada" & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    
    If objDataAccess.RunStoredProcedure(strSQL, gStringConnection) Then
        strSQL = "up_ModificarHolgura '" & iintSalida & "','" & "Salida" & "'"
        If objDataAccess.RunStoredProcedure(strSQL, gStringConnection) Then
            UpdateHolgura = True
        Else
            GoTo ExitMethod
        End If
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
    UpdateHolgura = False
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
     '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    UpdateHolgura = False
    
    Call RaiseError(vbObjectError + clsPersonal_UpdateHolgura, mstrClassModule & " - Update Holgura Method", Err.Description, CStr(Erl))
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
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetHolgura, mstrClassModule & " - Update Holgura Method", Err.Description, CStr(Erl))
End Function

'modificar para tolerancia

Public Function UpdateTolerancia(ByVal iintTolerancia As Integer) As Boolean
Dim objDataAccess As Object
Dim strSQL As String
On Error GoTo ErrHandler

'#If TRANSACTION_MODE = 1 Then
'    Dim ObjContexto As ObjectContext
'    Set ObjContexto = GetObjectContext
'#End If

    strSQL = "up_ModificarTolerancia '" & iintTolerancia & "','" & "Entrada" & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    
    If objDataAccess.RunStoredProcedure(strSQL, gStringConnection) Then
        strSQL = "up_ModificarTolerancia '" & iintTolerancia & "','" & "Tolerancia" & "'"
        If objDataAccess.RunStoredProcedure(strSQL, gStringConnection) Then
            UpdateTolerancia = True
        Else
            GoTo ExitMethod
        End If
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
    UpdateTolerancia = False
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
     '#If TRANSACTION_MODE = 1 Then
     '   ObjContexto.SetAbort
    '#End If
    UpdateTolerancia = False
    
    Call RaiseError(vbObjectError + clsPersonal_UpdateTolerancia, mstrClassModule & " - Update Tolerancia Method", Err.Description, CStr(Erl))
End Function
'fin de Tolerancia

Public Function GetCliente() As Object
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler

    strSQL = "SELECT Cliente, Tolerancia, Tiempo, ToleranciaTardanza, ToleranciaJornadaIncompleta, " & _
            " TiempoPrimeraMarcacion, HoraPrimeraMarcacion FROM Cliente "
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetCliente = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetCliente, mstrClassModule & " - GetCliente Method", Err.Description, CStr(Erl))
End Function

'*************************************
'esto es para un nuevo cliente osea deberia de ser cuando se instala por primera ves el sistema
Public Function AddNewCliente(ByVal istrCliente As String) As Boolean
    
Dim objDataAccess As Object
Dim intCodigo As Integer
Dim strSQL As String
    
On Error GoTo ErrHandler
    
    strSQL = "up_AgregarCliente '" & istrCliente & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    AddNewCliente = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    AddNewCliente = True
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewCliente = False
    Call RaiseError(vbObjectError + clsPersonal_AddNewCliente, mstrClassModule & " - AddNewCliente Method", Err.Description, CStr(Erl))
    'falta esto!!!!
End Function

'modificar Cliente
Public Function UpdateCliente(ByVal iintTiempo As Integer, _
                            ByVal istrCliente As String, _
                            ByVal iintTolerancia As Integer, _
                            ByVal iintTardanza As Integer, _
                            ByVal iintJorIncompleta As Integer, _
                            ByVal iintTiempoPrimeraMarcacion As Integer, _
                            ByVal idatHoraPrimeraMarcacion As Date) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    strSQL = "up_ModificarCliente '" & iintTiempo & "','" & istrCliente & "','" & iintTolerancia & "','" _
                & iintTardanza & "','" & iintJorIncompleta & "','" & iintTiempoPrimeraMarcacion & "','" _
                & FormatDateTime(idatHoraPrimeraMarcacion, vbShortTime) & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    UpdateCliente = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateCliente = False
    Call RaiseError(vbObjectError + clsPersonal_UpdateCliente, mstrClassModule & " - UpdateCliente Method", Err.Description, CStr(Erl))
End Function

'*************************************
'*** Encargado de mantener la integridad de los datos
Public Function ValidateBussines(ByVal iintAction As Integer, ByRef ostrMensajeError As String, _
                    Optional ByVal iintIdTrabajador As Integer = 0, _
                    Optional ByVal iintIdArea As Integer = 0, _
                    Optional ByVal iintIdCargo As Integer = 0, _
                    Optional ByVal iintIdGrupo As Integer = 0, _
                    Optional ByVal iintIdCategoria As Integer = 0, _
                    Optional ByVal iintIdTipoHora As Integer = 0, _
                    Optional ByVal iintIdMotivoJustificacion As Integer = 0) As Boolean

Dim objDataAccess As Object
Dim rstConsulta As Object
Dim strSQL As String
 
On Error GoTo ErrHandler
    Select Case iintAction
        
        '****************** Area *****************
        Case EnumActionsPersonal.Delete_Area
        
            strSQL = "up_BuscarTrabajador '" & 0 & "','" & iintIdArea & "','" & 0 & "','" & 0 & "','" & 0 & "'"
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, existen Trabajadores registrados en esta Área"
                GoTo ExitMethod
            End If
                        
        '****************** Cargo ******************
        Case EnumActionsPersonal.Delete_Cargo
            strSQL = "up_BuscarTrabajador '" & 0 & "','" & 0 & "','" & iintIdCargo & "','" & 0 & "','" & 0 & "'"
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, existen Trabajadores registrados en este Cargo"
                GoTo ExitMethod
            End If
            
        '****************** Categoria ******************
        Case EnumActionsPersonal.Delete_Categoria
            strSQL = "up_BuscarTrabajador '" & 0 & "','" & 0 & "','" & 0 & "','" & iintIdCategoria & "','" & 0 & "'"
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, existen Trabajadores registrados en esta Categoria"
                GoTo ExitMethod
            End If
        
        '****************** Grupo **********************
        Case EnumActionsPersonal.Delete_Grupo
            strSQL = "up_BuscarTrabajador '" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & iintIdGrupo & "'"
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, existen Trabajadores registrados en este Grupo"
                GoTo ExitMethod
            End If
        '************************* MotivoJustificacion **********************
        Case EnumActionsPersonal.Delete_MotivoJustificacion
            strSQL = "up_BuscarJustificacion '" & 0 & "','" & iintIdMotivoJustificacion & "'"
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, existen Justificaciones registradas con este Motivo de Justificacion"
                GoTo ExitMethod
            End If
        
            strSQL = "SELECT IdMotivoJustificacion FROM Justificacion WHERE IdMotivoJustificacion = " & gintMotivoFeriadoRemunerado & " OR IdMotivoJustificacion=" & gintMotivoFeriadoNoRemunerado
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If iintIdMotivoJustificacion = gintMotivoFeriadoRemunerado Or iintIdMotivoJustificacion = gintMotivoFeriadoNoRemunerado Or iintIdMotivoJustificacion = gintMotivoPermiso Then
                ostrMensajeError = "No es posible eliminar el registro, este registro es del Sistema"
                GoTo ExitMethod
            End If
        '************************* TipoHora ***************************
        Case EnumActionsPersonal.Delete_TipoHora
            strSQL = "up_BuscarHorarioxTrabajador '" & 0 & "','" & iintIdTipoHora & "','" & 0 & "'"
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, existen Horarios registrados con este Tipo de Hora"
                GoTo ExitMethod
            End If
        '************************** Trabajador ***************************
        Case EnumActionsPersonal.Delete_Trabajador
            strSQL = "up_BuscarHorarioxTrabajador '" & 0 & "','" & 0 & "','" & iintIdTrabajador & "'"
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, existen Horarios registrados a este Trabajador"
                GoTo ExitMethod
            End If
        
            strSQL = "up_BuscarDetalleConsolidadoAsistencia '" & iintIdTrabajador & "'"
            Set objDataAccess = CreateObject("Personal.clsData")
            Set rstConsulta = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
            Set objDataAccess = Nothing
            If Not rstConsulta.EOF Then
                ostrMensajeError = "No es posible eliminar el registro, existen Consolidados de Asistencia registrados a este Trabajador"
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

    Err.Raise Number:=vbObjectError + clsHorario_ValidateBussines, Source:=mstrClassModule & " - ValidateBussines Method", Description:=Err.Description
End Function


Public Function GetHorarioSemana(ByVal idblIdTRabajador As String, ByVal idatFecha As Date) As Object
Dim objDataAccess As Object
Dim strSQL As String
On Error GoTo ErrHandler

    strSQL = "up_BuscarHorarioSemana " & idblIdTRabajador & ", '" & idatFecha & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetHorarioSemana = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsPersonal_GetHolgura, mstrClassModule & " - Update Holgura Method", Err.Description, CStr(Erl))
End Function

