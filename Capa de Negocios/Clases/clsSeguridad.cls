VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const mstrClassModule As String = "Personal - cSeguridad Class Module"

Public Function AddNewUsuario(ByVal istrNombre As String, ByVal istrClave As String) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
10    strSQL = "up_AgregarUsuario '" & istrNombre & "','" & istrClave & "')"
20    Set objDataAccess = CreateObject("Personal.clsData")
30    AddNewUsuario = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
40    Set objDataAccess = Nothing
50    AddNewUsuario = True
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewUsuario = False
    Call RaiseError(vbObjectError + clsSeguridad_AddNewUsuario, mstrClassModule & " - AddNewUsuario Method", Err.Description, CStr(Erl))
    
End Function

Public Function UpdateUsuario(ByVal iintIdUsuario As Integer, ByVal istrNombre As String, _
                            ByVal istrClave As String) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler
    
10    strSQL = "up_ModificarUsuario '" & iintIdUsuario & "','" & istrNombre & "','" & istrClave & "'"
20    Set objDataAccess = CreateObject("Personal.clsData")
30    UpdateUsuario = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
40    Set objDataAccess = Nothing
50    UpdateUsuario = True
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateUsuario = False
    Call RaiseError(vbObjectError + clsSeguridad_UpdateUsuario, mstrClassModule & " - UpdateUsuario Method", Err.Description, CStr(Erl))
    
End Function

Public Function DeleteUsuario(ByVal iintIdUsuario As Integer, _
                        Optional ByRef ostrError As String = "") As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String
    
On Error GoTo ErrHandler

    If ValidateBussines(EnumActionsSeguridad.Delete_Usuario, ostrError, iintIdUsuario) Then
        
        strSQL = "up_EliminarUsuario '" & iintIdUsuario & "'"
        Set objDataAccess = CreateObject("Personal.clsData")
        DeleteUsuario = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
        Set objDataAccess = Nothing

    Else
        GoTo ExitMethod
    End If

Exit Function
ExitMethod:
    Set objDataAccess = Nothing
    DeleteUsuario = False
    
Exit Function
DeleteUsuario = True
ErrHandler:
    Set objDataAccess = Nothing
    DeleteUsuario = False
    Call RaiseError(vbObjectError + clsSeguridad_DeleteUsuario, mstrClassModule & " - DeleteUsuario Method", Err.Description, CStr(Erl))
    
End Function

Public Function GetUsuario(Optional ByVal iintIdUsuario As Integer = 0, _
                           Optional ByVal istrNombre As String = "", _
                           Optional ByVal istrNombreAnterior As String = "") As Object
    
Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
10  strSQL = "SELECT IdUsuario, NombreUsuario, Clave FROM Usuario WHERE 1=1 "
             
    If iintIdUsuario <> 0 Then
20      strSQL = strSQL & " AND IdUsuario = " & iintIdUsuario
    End If
    If istrNombre <> "" Then
30      strSQL = strSQL & " AND NombreUsuario = '" & istrNombre & "'"
    End If
    If istrNombreAnterior <> "" Then
        strSQL = strSQL & " AND NombreUsuario<>'" & istrNombreAnterior & "'"
    End If
40  Set objDataAccess = CreateObject("Personal.clsData")
50  Set GetUsuario = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
60  Set objDataAccess = Nothing
    
Exit Function
ErrHandler:
    Set GetUsuario = Nothing
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + clsSeguridad_GetUsuario, mstrClassModule & " - GetUsuario Method", Err.Description, CStr(Erl))

End Function

Public Function GetValidarUsuario(ByVal istrNombre As String, _
                           ByVal istrClave As String) As Object
    
Dim objDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
10  strSQL = "up_BuscarValidaUsuario '" & istrNombre & "','" & istrClave & "'"
20  Set objDataAccess = CreateObject("Personal.clsData")
30  Set GetValidarUsuario = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
50  Set objDataAccess = Nothing

Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set GetValidarUsuario = Nothing
    Call RaiseError(vbObjectError + clsSeguridad_GetValidarUsuario, mstrClassModule & " - GetUsuario Method", Err.Description, CStr(Erl))

End Function

'*************************************
'*** Encargado de mantener la integridad de los datos
Public Function ValidateBussines(ByVal iintAction As Integer, ByRef ostrMensajeError As String, _
                    Optional ByVal iintIdUsuario As Integer = 0) As Boolean

Dim objDataAccess As Object
Dim rstConsulta As Object
Dim strSQL As String
 
On Error GoTo ErrHandler
    Select Case iintAction
        
        '****************** Usuario***********************
        Case EnumActionsSeguridad.Delete_Usuario
            If iintIdUsuario = gintUsuarioAdministrador Then
                ostrMensajeError = "No es posible eliminar el usuario administrador"
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

    Err.Raise Number:=vbObjectError + clsSeguridad_ValidateBussines, Source:=mstrClassModule & " - ValidateBussines Method", Description:=Err.Description
End Function

Public Function GetAccesoUnicoUsuario(Optional ByVal iintIdUsuario As Integer = 0) As Object
Dim strSQL As String
Dim objDataAccess As Object

On Error GoTo ErrHandler

    strSQL = "up_BuscarAccesoUsuario '" & iintIdUsuario & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetAccesoUnicoUsuario = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    
Exit Function
ErrHandler:
    Set GetAccesoUnicoUsuario = Nothing
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsSeguridad_GetAccesoUnicoUsuario, Source:=mstrClassModule & " - GetMenuUsuario Method", Description:=Err.Description
End Function

Public Function GetOpcionesModulo(ByVal iintIdModulo As Integer) As Object
Dim strSQL As String
Dim objDataAccess As Object

On Error GoTo ErrHandler

    strSQL = "up_BuscarOpcionesModulo " & iintIdModulo
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetOpcionesModulo = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
    
Exit Function
ErrHandler:
    Set GetOpcionesModulo = Nothing
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsSeguridad_GetOpcionesModulo, Source:=mstrClassModule & " - GetMenuUsuario Method", Description:=Err.Description
End Function

Public Function GetOpcionMenuDeUsuario(ByVal iintIdUsuario As Integer) As Object
Dim strSQL As String
Dim objDataAccess As Object

On Error GoTo ErrHandler
    strSQL = "up_BuscarOpcionMenuDeUsuario '" & iintIdUsuario & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetOpcionMenuDeUsuario = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection, DataAccessLockOptimistic)
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Set GetOpcionMenuDeUsuario = Nothing
    Set objDataAccess = Nothing
    Err.Raise Number:=vbObjectError + clsSeguridad_GetOpcionMenuDeUsuario, Source:=mstrClassModule & " - GetMenuUsuario Method", Description:=Err.Description
End Function
 
Public Function UpdateAccesoUsuario(ByVal iintIdUsuario As Integer, ByVal iArrayMenu As Variant) As Boolean
    
Dim objDataAccess As Object
Dim intCodigo As Integer
Dim intContador As Integer
Dim strSQL As String
    
On Error GoTo ErrHandler

    Set objDataAccess = CreateObject("Personal.clsData")
    For intContador = LBound(iArrayMenu, 1) To UBound(iArrayMenu, 1)
        strSQL = "up_ModificarAccesoUsuario '" & iintIdUsuario & "','" & iArrayMenu(intContador, 0) & "','" & iArrayMenu(intContador, 2) & "'"
        UpdateAccesoUsuario = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Next
    
    Set objDataAccess = Nothing
    UpdateAccesoUsuario = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    UpdateAccesoUsuario = False
    Call RaiseError(vbObjectError + clsSeguridad_UpdateAccesoUsuario, mstrClassModule & " - AddNewJustificacion Method", Err.Description, CStr(Erl))
End Function

Public Function GetNombreCliente() As String
    
Dim objDataAccess As Object
Dim rstDataAccess As Object
Dim strSQL As String
   
On Error GoTo ErrHandler
 
  strSQL = "up_BuscarCliente"
  Set objDataAccess = CreateObject("Personal.clsData")
  Set rstDataAccess = objDataAccess.GetDataFromStoredProcedure(strSQL, gStringConnection)
  Set objDataAccess = Nothing
    GetNombreCliente = rstDataAccess!Cliente
    Set rstDataAccess = Nothing
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Set rstDataAccess = Nothing
    Call RaiseError(vbObjectError + clsSeguridad_GetNombreCliente, mstrClassModule & " - GetNombreCliente Method", Err.Description, CStr(Erl))

End Function
