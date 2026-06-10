VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCreaBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function CreaBase(ByVal istrSql As String) As Boolean
    
Dim objDataAccess As Object

On Error GoTo ErrHandler
    
    Set objDataAccess = CreateObject("Personal.clsData")
    CreaBase = objDataAccess.RunSQLSentence(istrSql, gStringConnectionCrear)
    Set objDataAccess = Nothing
    CreaBase = True
     
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    CreaBase = False
    Call RaiseError(vbObjectError + clsCreaBase_CreaBase, mstrClassModule & " - CreaBase Method", Err.Description, CStr(Erl))
End Function

Public Function Restaura(ByVal istrSP As String) As Boolean

    Dim cnx As ADODB.Connection
    
On Error GoTo ErrHandler

    Set cnx = CreateObject("ADODB.Connection")
    cnx.ConnectionString = gStringConnectionCrear
    cnx.Open
    cnx.Execute istrSP, , ADODB.adExecuteNoRecords
    cnx.Close
    Set cnx = Nothing
    Restaura = True
    
Exit Function
ErrHandler:
    Set cnx = Nothing
    Restaura = False
    Call RaiseError(vbObjectError + cData_RunSQLSentence, mstrClassModule & " - RunSQLSentence Method", Err.Description, CStr(Erl))
End Function

'Public Function Restaura(ByVal istrSql As String) As Boolean
'
'Dim objDataAccess As Object
'
'On Error GoTo ErrHandler
'
'    Set objDataAccess = CreateObject("Personal.clsData")
'    Restaura = objDataAccess.RunSQLCrear(istrSql, gStringConnectionCrear)
'    Set objDataAccess = Nothing
'    Restaura = True
'
'Exit Function
'ErrHandler:
'    Set objDataAccess = Nothing
'    Restaura = False
'    Call RaiseError(vbObjectError + clsCreaBase_Restaura, mstrClassModule & " - Restaura Method", Err.Description, CStr(Erl))
'End Function

Public Function BorrarObjetos(ByVal istrSql As String) As Boolean
   
Dim objDataAccess As Object

On Error GoTo ErrHandler
    
    Set objDataAccess = CreateObject("Personal.clsData")
    BorrarObjetos = objDataAccess.RunSQLSentence(istrSql, gStringConnection)
    Set objDataAccess = Nothing
    BorrarObjetos = True
     
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    BorrarObjetos = False
    Call RaiseError(vbObjectError + clsCreaBase_BorrarObjetos, mstrClassModule & " - BorrarObjetos Method", Err.Description, CStr(Erl))
End Function

Public Function CrearTablas(ByVal istrSql As String) As Boolean
   
Dim objDataAccess As Object

On Error GoTo ErrHandler
    
    Set objDataAccess = CreateObject("Personal.clsData")
    CrearTablas = objDataAccess.RunSQLSentence(istrSql, gStringConnection)
    Set objDataAccess = Nothing
    CrearTablas = True
     
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    CrearTablas = False
    Call RaiseError(vbObjectError + clsCreaBase_CrearTablas, mstrClassModule & " - CrearTablas Method", Err.Description, CStr(Erl))
End Function

Public Function CrearProcedimientos(ByVal istrSql As String) As Boolean
   
Dim objDataAccess As Object

On Error GoTo ErrHandler
    
    Set objDataAccess = CreateObject("Personal.clsData")
    CrearProcedimientos = objDataAccess.RunSQLSentence(istrSql, gStringConnection)
    Set objDataAccess = Nothing
    CrearProcedimientos = True
     
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    CrearProcedimientos = False
    Call RaiseError(vbObjectError + clsCreaBase_CrearProcedimientos, mstrClassModule & " - CrearProcedimientos Method", Err.Description, CStr(Erl))
End Function
