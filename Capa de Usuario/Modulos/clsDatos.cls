VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const gstrNameFile = "DataAccess.ini"
Const gstrSection = "Connection"
Const gstrKey = "StringConnection"

Public Enum enumDataAccessCursorType
    DataAccessOpenForwardOnly = 0
    DataAccessOpenStatic = 3
End Enum

Public Enum enumDataAccessLockType
    DataAccessLockReadOnly = 1
    DataAccessLockOptimistic = 3
    DataAccessLockUpdateBathOptimistic = 4
End Enum

Private gstrRegistryEntry As String
Private gstrKeyEncrypt As String
Private Const mstrClassModule As String = "Personal - cData Class Module"

Public Function RunSQLSentence(ByVal istrSP As String, _
                            ByVal istrConnection As String) As Boolean

    Dim cnx As ADODB.Connection
    
On Error GoTo ErrHandler
    
'    '#If TRANSACTION_MODE = 1 Then
'        Dim objContext As ObjectContext
'        Set objContext = GetObjectContext
'    '#End If
    
    Set cnx = New ADODB.Connection
    cnx.ConnectionString = istrConnection
    cnx.Open
    
    cnx.BeginTrans
    cnx.Execute istrSP, , ADODB.adExecuteNoRecords
    cnx.CommitTrans

    cnx.Close
    Set cnx = Nothing
    RunSQLSentence = True
    
'    '#If TRANSACTION_MODE = 1 Then
'        objContext.SetComplete
'    '#End If

Exit Function
ErrHandler:
    cnx.RollbackTrans
    Set cnx = Nothing
    RunSQLSentence = False
'    '#If TRANSACTION_MODE = 1 Then
'        objContext.SetAbort
'    '#End If
    
    'Call RaiseError(vbObjectError + cData_RunSQLSentence, mstrClassModule & " - RunSQLSentence Method", Err.Description, CStr(Erl))
End Function

Public Function RunSQLCrear(ByVal istrSP As String, _
                            ByVal istrConnection As String) As Boolean

    Dim cnx As ADODB.Connection
    Dim strCadenaConexion As String
    
On Error GoTo ErrHandler
'    strCadenaConexion = "Data Source= (local) ;" & _
'                        "Integrated Security=SSPI;" & _
'                        "Persist Security Info=False;" & _
'                        "Initial Catalog=master "
                        
    Set cnx = New ADODB.Connection
'    cnx.CursorLocation = adUseClient
'
'    cnx.Provider = "SQLOLEDB.1"
    cnx.ConnectionString = istrConnection
    cnx.Open
    cnx.BeginTrans
    cnx.Execute istrSP
    cnx.Close
    cnx.CommitTrans
    Set cnx = Nothing
    RunSQLCrear = True

Exit Function
ErrHandler:
    cnx.RollbackTrans
    Set cnx = Nothing
    RunSQLCrear = False
   
    'Call RaiseError(vbObjectError + "", mstrClassModule & " - RunSQLCrear Method", Err.Description, CStr(Erl))
End Function

Public Function GetDataFromSQLSentence(ByVal istrSP As String, _
                            Optional ByVal istrConnection As String = "", _
                            Optional ByVal iintLockType As enumDataAccessLockType = DataAccessLockReadOnly, _
                            Optional ByVal iintCursorType As enumDataAccessCursorType) As ADODB.Recordset

    Dim cnx As ADODB.Connection
    Dim rst As ADODB.Recordset
    
On Error GoTo ErrHandler
    
'    '#If TRANSACTION_MODE = 1 Then
'        Dim objContext As ObjectContext
'        Set objContext = GetObjectContext
'    '#End If
    
    Set cnx = New ADODB.Connection
    Set rst = New ADODB.Recordset
    cnx.ConnectionString = istrConnection
    cnx.Open
    
    rst.CursorLocation = adUseClient
    'rst.Open istrSP, cnx, iintCursorType, iintLockType
    rst.Open istrSP, cnx, adOpenKeyset, adLockOptimistic
    rst.ActiveConnection = Nothing
    
    cnx.Close
    
    Set GetDataFromSQLSentence = rst
    Set cnx = Nothing
    Set rst = Nothing
    
'    '#If TRANSACTION_MODE = 1 Then
'        objContext.SetComplete
'    '#End If
    
Exit Function
ErrHandler:
    Set cnx = Nothing
    Set rst = Nothing
    Set GetDataFromSQLSentence = Nothing
'    '#If TRANSACTION_MODE = 1 Then
'        objContext.SetAbort
'    '#End If
    'Call RaiseError(vbObjectError + cData_GetDataFromSQLSentence, mstrClassModule & " - GetDataFromSQLSentence Method", Err.Description, CStr(Erl))
End Function

Public Function RunStoredProcedure(ByVal istrSP As String, _
                                ByVal istrConnection As String) As Boolean
    
    Dim cnx As ADODB.Connection
    Dim cmd As ADODB.Command
    
On Error GoTo ErrHandler
    
'    '#If TRANSACTION_MODE = 1 Then
'        Dim objContext As ObjectContext
'        Set objContext = GetObjectContext
'    '#End If
    
    Set cnx = New ADODB.Connection
    Set cmd = New ADODB.Command
    
    cnx.ConnectionString = istrConnection
    'cnx.ConnectionTimeout = 10000
    cnx.CursorLocation = adUseClient
    cnx.Open
    cnx.BeginTrans
    cmd.ActiveConnection = cnx
    'cmd.ActiveConnection.ConnectionTimeout = 1000
    Dim iPos As Integer
    iPos = InStr(1, istrSP, "(")
    If iPos = 0 Then
        iPos = InStr(1, istrSP, " ")
        If iPos = 0 Then
            istrSP = istrSP '& "()"
        Else
            istrSP = Mid(istrSP, 1, iPos - 1) + "(" + Mid(istrSP, iPos + 1, Len(istrSP) - iPos) + ")"
        End If
    End If
    cmd.CommandText = istrSP
    cmd.CommandType = adCmdStoredProc
    'cmd.CommandTimeout = 100000
    cmd.Execute , , ADODB.adExecuteNoRecords
    Set cmd.ActiveConnection = Nothing
    cnx.CommitTrans
    cnx.Close
    
    Set cnx = Nothing
    Set cmd = Nothing
    RunStoredProcedure = True
    
'    '#If TRANSACTION_MODE = 1 Then
'        objContext.SetComplete
'    '#End If
    
Exit Function
ErrHandler:
    cnx.RollbackTrans
    Set cnx = Nothing
    Set cmd = Nothing
    RunStoredProcedure = False
'    '#If TRANSACTION_MODE = 1 Then
'        objContext.SetAbort
'    '#End If
    'Call RaiseError(vbObjectError + cData_RunStoredProcedure, mstrClassModule & " - RunStoredProcedure Method", Err.Description, CStr(Erl))
End Function

Public Function GetDataFromStoredProcedure(ByVal istrSP As String, _
                            ByVal istrConnection As String, _
                            Optional ByVal iintLockType As enumDataAccessLockType = DataAccessLockReadOnly, _
                            Optional ByVal iintCursorType As enumDataAccessCursorType) As ADODB.Recordset
                            
    Dim cnx As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
On Error GoTo ErrHandler
    
'    '#If TRANSACTION_MODE = 1 Then
'        Dim objContext As ObjectContext
'        Set objContext = GetObjectContext
'    '#End If
    
    'Set cnx = new ADODB.Connection
    Set cmd = New ADODB.Command
    Set rst = New ADODB.Recordset
    Set cnx = New ADODB.Connection
    
    cnx.ConnectionString = istrConnection
    cnx.Open
    cnx.CommandTimeout = 1000
    cnx.BeginTrans
    
    cmd.ActiveConnection = cnx
    cmd.CommandTimeout = 1000
    Dim iPos As Integer
    iPos = InStr(1, istrSP, "(")
    If iPos = 0 Then
        iPos = InStr(1, istrSP, " ")
        If iPos = 0 Then
            istrSP = istrSP '& "()"
        Else
            istrSP = Mid(istrSP, 1, iPos - 1) + "(" + Mid(istrSP, iPos + 1, Len(istrSP) - iPos) + ")"
        End If
    End If
    cmd.CommandText = istrSP
    cmd.CommandType = adCmdStoredProc
    
    rst.CursorLocation = adUseClient
    'rst.Open cmd, , iintCursorType, iintLockType
    rst.Open cmd, , adOpenKeyset, adLockOptimistic
    Set rst.ActiveConnection = Nothing
    cnx.CommitTrans
    cnx.Close
    
    Set GetDataFromStoredProcedure = rst
    Set cnx = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    
'    '#If TRANSACTION_MODE = 1 Then
'        objContext.SetComplete
'    '#End If
    
Exit Function
ErrHandler:
    cnx.RollbackTrans
    Set cnx = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Set GetDataFromStoredProcedure = Nothing
    
'    '#If TRANSACTION_MODE = 1 Then
'        objContext.SetAbort
'    '#End If
    
    'Call RaiseError(vbObjectError + cData_GetDataFromStoredProcedure, mstrClassModule & " - GetDataFromStoredProcedur Method", Err.Description, CStr(Erl))
End Function

Public Function RunSQLComplex(ByVal istrSQLCabecera As String, ByVal iArraySQLDetalle As Variant, ByVal istrConnection As String) As Boolean
    
Dim lngRowsAffected As Long
Dim intContador As Integer
Dim cnx As ADODB.Connection

On Error GoTo ErrHandler
    Set cnx = New ADODB.Connection
    cnx.ConnectionString = gstrConexion

    cnx.Open
    
    cnx.BeginTrans
    cnx.IsolationLevel = adXactSerializable
    cnx.Execute istrSQLCabecera, lngRowsAffected, ADODB.adExecuteNoRecords
    
    For intContador = LBound(iArraySQLDetalle, 1) To UBound(iArraySQLDetalle, 1) - 1
        If iArraySQLDetalle(intContador) = "" Then
            Exit For
        End If
        cnx.Execute iArraySQLDetalle(intContador), lngRowsAffected, ADODB.adExecuteNoRecords
    Next
    cnx.CommitTrans

    RunSQLComplex = True
    Set cnx = Nothing
Exit Function
ErrHandler:
    cnx.RollbackTrans
    Set cnx = Nothing
    RunSQLComplex = False
    'Call RaiseError(vbObjectError + cData_RunSQLSentence, mstrClassModule & " - RunSQLSentence Method", Err.Description, CStr(Erl))
End Function

Public Function RunSQLSoloDetalle(ByVal iArraySQLDetalle As Variant, ByVal istrConnection As String) As Boolean
    
Dim lngRowsAffected As Long
Dim intContador As Integer
Dim cnx As ADODB.Connection

On Error GoTo ErrHandler
    Set cnx = New ADODB.Connection

    cnx.ConnectionString = istrConnection
    cnx.Open
    cnx.BeginTrans
    cnx.IsolationLevel = adXactSerializable
    For intContador = LBound(iArraySQLDetalle, 1) To UBound(iArraySQLDetalle, 1) - 1
        If iArraySQLDetalle(intContador) = "" Then
            Exit For
        End If
        cnx.Execute iArraySQLDetalle(intContador), lngRowsAffected, ADODB.adExecuteNoRecords
    Next
    
    cnx.CommitTrans
    RunSQLSoloDetalle = True
    Set cnx = Nothing
Exit Function
ErrHandler:
    cnx.RollbackTrans
    Set cnx = Nothing
    RunSQLSoloDetalle = False
    'Call RaiseError(vbObjectError + cData_RunSQLSoloDetalle, mstrClassModule & " - RunSQLSoloDetalle Method", Err.Description, CStr(Erl))

End Function

Public Function RunStoredProcedureSoloDetalle(ByVal iArraySQLDetalle As Variant, ByVal istrConnection As String) As Boolean
    
Dim lngRowsAffected As Long
Dim intContador As Integer
Dim cnx As ADODB.Connection
Dim cmd As ADODB.Command

On Error GoTo ErrHandler
    
'    '#If TRANSACTION_MODE = 1 Then
'        Dim objContext As ObjectContext
'        Set objContext = GetObjectContext
'    '#End If
    Dim iPos As Integer
    Dim istrSP As String
    Set cnx = New ADODB.Connection
    Set cmd = New ADODB.Command
    
    cnx.ConnectionString = istrConnection
    cnx.Open
    
    cnx.BeginTrans
    
    cnx.IsolationLevel = adXactSerializable
    cmd.ActiveConnection = cnx
    'cmd.CommandText = istrSP
    cmd.CommandType = adCmdStoredProc
    
    For intContador = LBound(iArraySQLDetalle, 1) To UBound(iArraySQLDetalle, 1) - 1
'        If iArraySQLDetalle(intContador) = "" Then
'            Exit For
'        End If
'        cnx.Execute iArraySQLDetalle(intContador), lngRowsAffected, ADODB.adExecuteNoRecords
    
        If iArraySQLDetalle(intContador) <> "" Then
            istrSP = iArraySQLDetalle(intContador)
            iPos = InStr(1, istrSP, "(")
            If iPos = 0 Then
                iPos = InStr(1, istrSP, " ")
                If iPos = 0 Then
                    istrSP = istrSP '& "()"
                Else
                    istrSP = Mid(istrSP, 1, iPos - 1) + "(" + Mid(istrSP, iPos + 1, Len(istrSP) - iPos) + ")"
                End If
            End If
            cmd.CommandText = istrSP
            cmd.Execute , , ADODB.adExecuteNoRecords
        End If
    Next

    Set cmd.ActiveConnection = Nothing
    cnx.CommitTrans
    
    cnx.Close
    
    Set cnx = Nothing
    Set cmd = Nothing
    
    RunStoredProcedureSoloDetalle = True
    
'    '#If TRANSACTION_MODE = 1 Then
'        objContext.SetComplete
'    '#End If
    
Exit Function
ErrHandler:
    cnx.RollbackTrans
    Set cnx = Nothing
    Set cmd = Nothing
    RunStoredProcedureSoloDetalle = False
    
'    '#If TRANSACTION_MODE = 1 Then
'        objContext.SetAbort
'    '#End If
    'Call RaiseError(vbObjectError + cData_RunStoredProcedureSoloDetalle, mstrClassModule & " - RunSQLSoloDetalle Method", Err.Description, CStr(Erl))

End Function
