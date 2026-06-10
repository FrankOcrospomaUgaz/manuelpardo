Attribute VB_Name = "APIMgm"
Option Explicit
' API de manipulación del row (32 bits)
Declare Function OSRegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Declare Function OSRegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal LCszSubKey As String, phkResult As Long) As Long
Declare Function OSRegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal LCszSubKey As String) As Long
Declare Function OSRegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal LCszSubKey As String, phkResult As Long) As Long
Declare Function OSRegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal LCszValueName As String, ByVal dwReserved As Long, LCdwType As Long, LCbData As Any, cbData As Long) As Long
Declare Function OSRegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal LCszValueName As String, ByVal dwReserved As Long, ByVal fdwType As Long, LCbData As Any, ByVal cbData As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Public Enum EDirectorio
cHKEY_CLASSES_ROOT = &H80000000
cHKEY_CURRENT_USER = &H80000001
cHKEY_LOCAL_MACHINE = &H80000002
End Enum
Const ERROR_SUCCESS = 0&
Const Dat1_SZ = 1
Const Dat1_BINARY = 3
Const Dat1_DWORD = 4
' Crea (o abre si ya existe) una clave en el row del sistema
Function RegCreateKey(ByVal hKey As Long, ByVal LCszKey As String, phkResult As Long) As Boolean
    On Error GoTo 0
    If OSRegCreateKey(hKey, LCszKey, phkResult) = ERROR_SUCCESS Then
        RegCreateKey = True
    Else
        RegCreateKey = False
    End If
End Function
' Asocia un valor cnnSistema nombre (StrValueName = nombre) o sin nombre (StrValueName = "")
' cnnSistema una clave del row.
Function RegSetStringValue(ByVal hKey As Long, ByVal strValueName As String, ByVal StrData As String) As Boolean
    On Error GoTo 0
    If hKey = 0 Then Exit Function
    If OSRegSetValueEx(hKey, strValueName, 0&, Dat1_SZ, ByVal StrData, _
        Len(StrData) + 1) = ERROR_SUCCESS Then
        RegSetStringValue = True
    Else
        RegSetStringValue = False
    End If
End Function
' Asocia un valor cnnSistema nombre (strValueName = nombre) o sin nombre (strValueName = "")
'   cnnSistema una clave del row.
Function RegSetNumericValue(ByVal hKey As Long, ByVal strValueName As String, ByVal lData As Long, Optional ByVal fLog) As Boolean
    On Error GoTo 0
    If OSRegSetValueEx(hKey, strValueName, 0&, Dat1_DWORD, lData, 4) = ERROR_SUCCESS Then
        RegSetNumericValue = True
    Else
        RegSetNumericValue = False
    End If
End Function
' Abre una clave existente en el row del sistema.
Function RegOpenKey(ByVal hKey As Long, ByVal LCszSubKey As String, phkResult As Long) As Boolean
    On Error GoTo 0
    If OSRegOpenKey(hKey, LCszSubKey, phkResult) = ERROR_SUCCESS Then
        RegOpenKey = True
    Else
        RegOpenKey = False
    End If
End Function
' Elimina una clave existente del row del sistema.
Function RegDeleteKey(ByVal hKey As Long, ByVal LCszSubKey As String) As Boolean
    On Error GoTo 0
    If OSRegDeleteKey(hKey, LCszSubKey) = ERROR_SUCCESS Then
        RegDeleteKey = True
    Else
        RegDeleteKey = False
    End If
End Function
' Cierra una clave abierta del row
Function RegCloseKey(ByVal hKey As Long) As Boolean
    On Error GoTo 0
    If OSRegCloseKey(hKey) = ERROR_SUCCESS Then
        RegCloseKey = True
    Else
        RegCloseKey = False
    End If
End Function
' Recupera los datos de cadena para un valor cnnSistema nombre
' (StrValueName = nombre) o sin nombre (StrValueName = "")
' dentro de una clave del row. Si el valor cnnSistema
' nombre existe, pero sus datos no son una cadena, esta
' función fallará.
Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String, StrData As String) As Boolean
    On Error GoTo 0
    Dim lValueType As Long
    Dim StrBuf As String
    Dim lDataBufSize As Long
    RegQueryStringValue = False
    ' Obtiene el tipo y longitud de los datos
    If OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize) = ERROR_SUCCESS Then
        If lValueType = Dat1_SZ Then
            StrBuf = String(lDataBufSize, " ")
            If OSRegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal StrBuf, lDataBufSize) = ERROR_SUCCESS Then
                StrData = Left(StrBuf, lDataBufSize - 1)
                RegQueryStringValue = True
            End If
        End If
    End If
End Function
' Recupera los datos enteros para un valor cnnSistema nombre
' (StrValueName = nombre) o sin nombre (StrValueName = "")
' dentro de una clave del row. Si el valor cnnSistema nombre
' existe, pero sus datos no son de tipo Dat1_DWORD, esta
' función fallará.
Function RegQueryNumericValue(ByVal hKey As Long, ByVal strValueName As String, lData As Long) As Boolean
    On Error GoTo 0
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    RegQueryNumericValue = False
    ' Obtiene el tipo y longitud de los datos
    lDataBufSize = 4
    If OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, lBuf, lDataBufSize) = ERROR_SUCCESS Then
        If lValueType = Dat1_DWORD Then
            lData = lBuf
            RegQueryNumericValue = True
        End If
    End If
End Function
Function GrabarClave(ByVal Ubicacion As String, ByVal Clave As String, ByVal Valor As String, ByVal Directorio As EDirectorio) As Long
Dim hKey As Long
Dim hkeyExistente As Long
Dim hkeyNueva As Long
' Abrimos la clave del row
GrabarClave = 1
If RegOpenKey(Directorio, Ubicacion, hkeyExistente) Then
    hKey = hkeyExistente
Else
   ' Como no existe, la creamos
    If RegCreateKey(Directorio, Ubicacion, hkeyNueva) Then
    hKey = hkeyNueva
    Else
    GrabarClave = 0
    Exit Function
    End If
End If
' Establecer valor
 RegSetStringValue hKey, Clave, Valor
    
RegCloseKey hKey
End Function
Function LeerClave(ByVal Ubicacion As String, ByVal Clave As String, ByVal Directorio As EDirectorio) As String
 ' Abrimos la clave del row
Dim hkeyExistente As Long
Dim hkeyNueva As Long
Dim StrValor As String
Dim hKey As Long
If RegOpenKey(Directorio, Ubicacion, hkeyExistente) Then
    hKey = hkeyExistente
    If RegQueryStringValue(hkeyExistente, Clave, StrValor) Then
      LeerClave = StrValor
    Else
      LeerClave = ""
      Exit Function
    End If
Else
   ' Como no existe, la creamos
    LeerClave = ""
    If RegCreateKey(Directorio, Ubicacion, hkeyNueva) Then
        hKey = hkeyNueva
    Else
        Exit Function
    End If

End If
RegCloseKey hKey
End Function
Function getKeyForm(ByVal frm_name As String, ByVal key_name As String) As String
getKeyForm = APIMgm.LeerClave("SOFTWARE\" & "System" & "\" & frm_name, key_name, cHKEY_LOCAL_MACHINE)
End Function
Sub SaveKeyForm(ByVal frm_name As String, ByVal key_name As String, ByVal key_value As String)
APIMgm.GrabarClave "SOFTWARE\" & "System" & "\" & frm_name, key_name, key_value, cHKEY_LOCAL_MACHINE
End Sub
