VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mstrClassModule As String = "Personal - cGeneral Class Module"

Public Function GetCampos(ByVal istrCampos As String) As Object

Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler
        
    strSQL = "SELECT " & istrCampos & " FROM Campos"
    
    Set objDataAccess = CreateObject("Personal.clsData")
    Set GetCampos = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection, DataAccessLockUpdateBathOptimistic, DataAccessOpenForwardOnly)
    Set objDataAccess = Nothing
    
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    
    Call RaiseError(vbObjectError + clsGeneral_GetCampos, mstrClassModule & " - GetCampos Method", Err.Description, CStr(Erl))
End Function

Public Function AddNewBitacora(ByVal istrUsuario As String, _
                                ByVal istrTabla As String, _
                                ByVal istrOperacion As String, _
                                ByVal istrRegistro As String) As Boolean
    
Dim objDataAccess As Object
Dim strSQL As String

On Error GoTo ErrHandler
 
    strSQL = "up_AgregarBitacora '" & istrUsuario & "','" & istrTabla & "','" & istrOperacion & "','" & istrRegistro & "'"
    Set objDataAccess = CreateObject("Personal.clsData")
    AddNewBitacora = objDataAccess.RunStoredProcedure(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    AddNewBitacora = True
     
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    AddNewBitacora = False
    Call RaiseError(vbObjectError + clsGeneral_AddNewBitacora, mstrClassModule & " - AddNewBitacora Method", Err.Description, CStr(Erl))
End Function

