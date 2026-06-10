VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
    DataAccessLockUpdateBathOptimistic = 4
End Enum

Private gstrRegistryEntry As String
Private gstrKeyEncrypt As String
Private Const mstrClassModule As String = "Personal - cData Class Module"

Function GetStringConnection() As String
  gstrKeyEncrypt = IIf(gstrKeyEncrypt = "", "FASTCOM PERU", gstrKeyEncrypt)
  GetStringConnection = Decrypt(ReadFileIni(gstrNameFile, gstrSection, gstrKey), IIf(gstrKeyEncrypt = "", "FASTCOM PERU", gstrKeyEncrypt))
End Function

Public Function RunSQLSentence(ByVal istrSP As String, ByVal istrConnection As String) As Boolean
    
    Dim cnx As ADODB.Connection
    Dim lngRowsAffected As Long
    
On Error GoTo ErrHandler
    
    Set cnx = CreateObject("ADODB.Connection")
    If istrConnection = "" Then
        cnx.ConnectionString = GetStringConnection()
    Else
        cnx.ConnectionString = istrConnection
    End If
    cnx.Open
    cnx.BeginTrans
    cnx.Execute istrSP, lngRowsAffected, ADODB.adExecuteNoRecords
    cnx.CommitTrans
    cnx.Close
    Set cnx = Nothing
    RunSQLSentence = True
Exit Function
ErrHandler:
    cnx.RollbackTrans
    Set cnx = Nothing
    RunSQLSentence = False
    Call RaiseError(vbObjectError + cData_RunSQLSentence, mstrClassModule & " - RunSQLSentence Method", Err.Description, CStr(Erl))
End Function

Public Function GetDataFromSQLSentence(ByVal istrSP As String, _
                            Optional ByVal istrConnection As String = "", _
                            Optional ByVal iintLockType As enumDataAccessLockType = DataAccessLockReadOnly, _
                            Optional ByVal iintCursorType As enumDataAccessCursorType = DataAccessOpenForwardOnly) As ADODB.Recordset

    
    Dim cnx As ADODB.Connection
    Dim rst As ADODB.Recordset
    
On Error GoTo ErrHandler
       
    Set cnx = CreateObject("ADODB.Connection")
    Set rst = CreateObject("ADODB.Recordset")
    
    If istrConnection = "" Then
        cnx.ConnectionString = GetStringConnection()
    Else
        cnx.ConnectionString = istrConnection
    End If
    cnx.Open
    
    rst.CursorLocation = adUseClient
    rst.Open istrSP, cnx, iintCursorType, iintLockType
    rst.ActiveConnection = Nothing
    
    cnx.Close
    
    Set GetDataFromSQLSentence = rst
    Set cnx = Nothing
    Set rst = Nothing
    
Exit Function
ErrHandler:
    Set cnx = Nothing
    Set rst = Nothing
    Call RaiseError(vbObjectError + cData_GetDataFromSQLSentence, mstrClassModule & " - GetDataFromSQLSentence Method", Err.Description, CStr(Erl))
End Function

Public Function RunSQLComplex(ByVal istrSQLCabecera As String, ByVal iArraySQLDetalle As Variant, ByVal istrConnection As String) As Boolean
    
    Dim cnx As ADODB.Connection
    Dim lngRowsAffected As Long
    Dim intContador As Integer
    
On Error GoTo ErrHandler
    
    Set cnx = CreateObject("ADODB.Connection")
    If istrConnection = "" Then
        cnx.ConnectionString = GetStringConnection()
    Else
        cnx.ConnectionString = istrConnection
    End If
    cnx.Open
    cnx.BeginTrans
    cnx.Execute istrSQLCabecera, lngRowsAffected, ADODB.adExecuteNoRecords
    
    For intContador = LBound(iArraySQLDetalle, 1) To UBound(iArraySQLDetalle, 1) - 1
        If iArraySQLDetalle(intContador) = "" Then
            Exit For
        End If
        cnx.Execute iArraySQLDetalle(intContador), lngRowsAffected, ADODB.adExecuteNoRecords
    Next
    cnx.CommitTrans
    cnx.Close
    Set cnx = Nothing

    RunSQLComplex = True
Exit Function
ErrHandler:
    cnx.RollbackTrans
    Set cnx = Nothing
    RunSQLComplex = False
    Call RaiseError(vbObjectError + cData_RunSQLSentence, mstrClassModule & " - RunSQLSentence Method", Err.Description, CStr(Erl))
End Function

Public Function RunSQLSoloDetalle(ByVal iArraySQLDetalle As Variant, ByVal istrConnection As String) As Boolean
    
    Dim cnx As ADODB.Connection
    Dim lngRowsAffected As Long
    Dim intContador As Integer
    
On Error GoTo ErrHandler
    
    Set cnx = CreateObject("ADODB.Connection")
    If istrConnection = "" Then
        cnx.ConnectionString = GetStringConnection()
    Else
        cnx.ConnectionString = istrConnection
    End If
    
    cnx.Open
    cnx.BeginTrans
    
    For intContador = LBound(iArraySQLDetalle, 1) To UBound(iArraySQLDetalle, 1) - 1
        If iArraySQLDetalle(intContador) = "" Then
            Exit For
        End If
        cnx.Execute iArraySQLDetalle(intContador), lngRowsAffected, ADODB.adExecuteNoRecords
    Next
    
    cnx.CommitTrans
    cnx.Close
    Set cnx = Nothing
    RunSQLSoloDetalle = True

Exit Function
ErrHandler:
    cnx.RollbackTrans
    Set cnx = Nothing
    RunSQLSoloDetalle = False
    Call RaiseError(vbObjectError + cData_RunSQLSoloDetalle, mstrClassModule & " - RunSQLSoloDetalle Method", Err.Description, CStr(Erl))

End Function

Public Function WriteStringConnection(ByVal istrConnection As String) As Boolean
  gstrKeyEncrypt = IIf(gstrKeyEncrypt = "", "FASTCOM PERU", gstrKeyEncrypt)
  WriteStringConnection = WriteFileIni(gstrNameFile, gstrSection, gstrKey, Encrypt(istrConnection, IIf(gstrKeyEncrypt = "", "FASTCOM PERU", gstrKeyEncrypt)))
End Function







