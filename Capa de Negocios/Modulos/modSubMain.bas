Attribute VB_Name = "modSubMain"
 '*********** File API functions
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'Set to True to make the password case-sensitive
Const CASE_SENSITIVE_PASSWORD = False

'Public cnx As ADODB.Connection
Public gStringConnection As String
Public gStringConnectionCrear As String


Const gstrNameFile = "DataAccess.ini"
Const gstrSection = "Connection"
Const gstrKey = "StringConnection"

Private gstrRegistryEntry As String
Public Const gstrKeyEncrypt As String = "FASTCOM PERU"
Public Sub Main()
   ' gStringConnectionCrear = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=master;Data Source=PERSONAL"
    If gStringConnection = "" Then
        gStringConnection = GetStringConnection()
    End If
End Sub
Function GetStringConnection() As String
Dim istrB As String
Dim istrA As String
'  gstrKeyEncrypt = IIf(gstrKeyEncrypt = "", "FASTCOM PERU", gstrKeyEncrypt)
 ' GetStringConnection = Decrypt(ReadFileIni(gstrNameFile, gstrSection, gstrKey), IIf(gstrKeyEncrypt = "", "FASTCOM PERU", gstrKeyEncrypt))
 'GetStringConnection = "Driver={SQL Server};Server=Server;Database=Personal;"
 'ultima 'GetStringConnection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Personal;Data Source=(local)"
 istrB = "12JULIOS"
 istrA = "sa"
' GetStringConnection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Congreso;Data Source=(local)"
 'GetStringConnection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Congreso;Data Source=ADRED"
 'GetStringConnection = "Provider=SQLOLEDB.1;Password=" + istrB + ";Persist Security Info=True;User ID=" + istrA + ";Initial Catalog=Congreso;Data Source=PERSONAL"
 
  'GetStringConnection = "Provider=SQLOLEDB.1;Password=" + istrB + ";Persist Security Info=True;User ID=" + istrA + ";Initial Catalog=Congreso;Data Source=PERSONAL\SQLEXPRESS"
  GetStringConnection = "Provider=SQLOLEDB.1;Password=1209;Persist Security Info=True;User ID=personal;Initial Catalog=PERSONAL;Data Source=PERSONAL2\SQLEXPRESS"
 
 'GetStringConnection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Congreso;Data Source=PERSONAL"
End Function

Public Sub RaiseError(ByVal istrNumber As Long, ByVal istrSource As String, ByVal istrErrDescription As String, ByVal istrLine As String)
    App.LogEvent "Number: " & CStr(istrNumber) & " - Source: " & istrSource & " - Description: " & istrErrDescription & " - Line: " & istrLine, vbLogEventTypeError
    Err.Raise istrNumber, istrSource, istrErrDescription & " - Line: " & istrLine
End Sub

'Encrypt text
Public Function Encrypt(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

End If

    'Encrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    Encrypt = strBuff
End Function

'Decrypt text encrypted with EncryptText
Public Function Decrypt(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

End If

    'Decrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
   End If
    Decrypt = strBuff
End Function

Public Function ReadFileIni(ByVal strNameFile As String, _
                            ByVal strSection As String, _
                            ByVal strKey As String) As String
  
  Dim intPartLen As Integer
  Dim strPart As String
  strPart = Space(255)
  intPartLen = GetPrivateProfileString(strSection, strKey, "", strPart, Len(strPart), App.Path & "\" & strNameFile)
  strPart = Left(strPart, intPartLen)
  ReadFileIni = strPart
End Function


Public Function WriteFileIni(ByVal strNameFile As String, _
                             ByVal strSection As String, _
                             ByVal strKey As String, _
                             ByVal strValue As String)
                            
    WritePrivateProfileString strSection, strKey, strValue, App.Path & "\" & strNameFile
End Function


Public Function UpdateSecuenciaCodigo(Optional ByVal istrTabla As String = "", _
                                      Optional ByVal iintAumento As Integer = 1) As Boolean
    
    Dim objDataAccess As Object
    Dim strSQL As String
    Dim blnResult As Boolean

 On Error GoTo ErrHandler
    
    strSQL = "Update SecuenciaCodificacion Set Codigo=Codigo+" & iintAumento & " where Tabla = '" & istrTabla & "'"
    
    Set objDataAccess = CreateObject("Personal.clsData")
    blnResult = objDataAccess.RunSQLSentence(strSQL, gStringConnection)
    Set objDataAccess = Nothing
    UpdateSecuenciaCodigo = True
Exit Function
ErrHandler:
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + CGeneral_UpdateSecuenciaCodigo, " UpdateSecuenciaCodigo Method", Err.Description, CStr(Erl))
End Function

Public Function GetSecuenciaCodigo(Optional ByVal istrTabla As String = "") As Long
    
    Dim objDataAccess As Object
    Dim rstSecuencia As Object
    Dim strSQL As String
    Dim blnResult As Boolean

 On Error GoTo ErrHandler
    GetSecuenciaCodigo = 0
    strSQL = "Select Codigo From SecuenciaCodificacion where Tabla = '" & istrTabla & "'"
    
    Set objDataAccess = CreateObject("Personal.clsData")
    Set rstSecuencia = objDataAccess.GetDataFromSQLSentence(strSQL, gStringConnection)
    If rstSecuencia.RecordCount > 0 Then
        rstSecuencia.MoveFirst
        GetSecuenciaCodigo = rstSecuencia!Codigo
    End If
    Set rstSecuencia = Nothing
    Set objDataAccess = Nothing
Exit Function
ErrHandler:
    Set rstSecuencia = Nothing
    Set objDataAccess = Nothing
    Call RaiseError(vbObjectError + CGeneral_GetSecuenciaCodigo, " GetSecuenciaCodigo Method", Err.Description, CStr(Erl))
End Function

Public Sub RecToArray(ByVal rstarray As Object, ByRef ArrayDatos As Variant)
Dim intContador As Integer
Dim intColumnas As Integer
If rstarray.RecordCount > 0 Then
    rstarray.MoveFirst
    For intContador = 0 To UBound(ArrayDatos, 1)
        If rstarray.EOF = False Then
            For intColumnas = 0 To rstarray.Fields.Count - 1
                If IsNull(rstarray.Fields(intColumnas).Value) Then
                    ArrayDatos(intContador, intColumnas) = 0
                Else
                    ArrayDatos(intContador, intColumnas) = rstarray.Fields(intColumnas).Value
                End If
            Next
            rstarray.MoveNext
        Else
            Exit Sub
        End If
    Next
End If
End Sub

Public Function ObtenerMes(ByVal idatFecha As Date) As String
Dim Mes As Integer
    Mes = Month(idatFecha)
    Select Case Mes
        Case 1
            ObtenerMes = "ENERO"
        Case 2
            ObtenerMes = "FEBRERO"
        Case 3
            ObtenerMes = "MARZO"
        Case 4
            ObtenerMes = "ABRIL"
        Case 5
            ObtenerMes = "MAYO"
        Case 6
            ObtenerMes = "JUNIO"
        Case 7
            ObtenerMes = "JULIO"
        Case 8
            ObtenerMes = "AGOSTO"
        Case 9
            ObtenerMes = "SETIEMBRE"
        Case 10
            ObtenerMes = "OCTUBRE"
        Case 11
            ObtenerMes = "NOVIEMBRE"
        Case 12
            ObtenerMes = "DICIEMBRE"
    End Select
End Function


