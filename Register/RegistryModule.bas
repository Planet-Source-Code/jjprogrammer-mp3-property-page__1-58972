Attribute VB_Name = "RegistryModule"
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Const REG_BINARY = 3            ' Binary data
Public Const REG_DWORD = 4            ' 32-Bit long
Public Const REG_DWORD_BIG_ENDIAN = 5        ' 32-bit long (big endian)
Public Const REG_DWORD_LITTLE_ENDIAN = 4    ' 32-bit long (little endian)
Public Const REG_EXPAND_SZ = 2        ' Unexpanded environment string
Public Const REG_LINK = 6            ' Symbolic link
Public Const REG_MULTI_SZ = 7            ' list of strings
Public Const REG_NONE = 0            ' Undefined
Public Const REG_RESOURCE_LIST = 8        ' Resource list (for device drivers)
Public Const REG_SZ = 1            ' String with null terminator

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long

Declare Function RegCreateKey Lib "advapi32.dll" Alias _
        "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, _
        phkResult As Long) As Long

Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
        "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long

Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
        "RegDeleteValueA" (ByVal Hkey As Long, _
        ByVal lpValueName As String) As Long

Declare Function RegOpenKey Lib "advapi32.dll" Alias _
        "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, _
        phkResult As Long) As Long

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal Hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
        "RegQueryValueExA" (ByVal Hkey As Long, _
        ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, lpcbData As Long) As Long

Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
        "RegSetValueExA" (ByVal Hkey As Long, _
        ByVal lpValueName As String, ByVal Reserved As Long, _
        ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Function CreateFileAssociation(sAppExtension As String, sApplicationPath As String, sDescription As String) As Boolean
'Purpose     :  Creates a file association for a give file extension.
'Inputs      :  sAppExtension                   The file extension to associate.
'               sApplicationPath                The name of the file to open the specified files with.
'               sDescription                    The description of the file type eg. "Excel Workbook".
'               sIconPath                       The path to the file where the icon is stored.
'               [sIconIndex]                    The index of the icon within the path. If not specified
'                                               uses the first icon.
'Outputs     :  Returns True on success
'Author      :  Andrew Baker
'Date        :  30/01/2001 11:29
'Notes       :  If updating an existing value, you may need to restart the computer before the
'               changes take effect.
'               Example usage:
'               bResult = CreateFileAssociation(".txt", "notepad.exe", "A Notepad File")
'Revisions   :
    
    Dim Result&, bResult As Boolean, sKeyName As String
    Const HKEY_CLASSES_ROOT = &H80000000
    
    If Len(sIconPath) = 0 Then
        'Use the application file for the icon
        sIconPath = "c:\vb5\samples\audio\Johnson Media Player\"
    End If
    'Write associations into registry
    sKeyName = Right$(sAppExtension, 3) & " file"
    If Not GetString(HKEY_CLASSES_ROOT, sAppExtension, "") = sKeyName Then 'if not registered
        bResult = zRegistryCreateKey(HKEY_CLASSES_ROOT, sAppExtension, , sKeyName) 'create extension key
        bResult = zRegistryCreateKey(HKEY_CLASSES_ROOT, sAppExtension, "Content Type", sDescription) 'create extension key description (content type)
    End If
    'create file type key description
    If sDescription > "" Then bResult = bResult And zRegistryCreateKey(HKEY_CLASSES_ROOT, sKeyName, , sDescription)
    CreateFileAssociation = bResult
End Function
Public Function AssociateIcon(sKeyName As String, ByVal sIconPath As String, Optional sIconIndex As String = ",0") As Boolean
    Dim bResult As Boolean
    'create file type key default icon
    If sIconPath > "" Then bResult = bResult And zRegistryCreateKey(HKEY_CLASSES_ROOT, sKeyName & "\DefaultIcon", , sIconPath & sIconIndex)
    'Update the Windows Icon Cache to see updated icon right away:
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
    AssociateIcon = bResult
End Function
Public Function CreateExplorerShellCommand(sKeyName As String, sCommandName As String, sCommandString As String) As Boolean
Dim bResult As Boolean
    'create file type key shell commands:
    bResult = zRegistryCreateKey(HKEY_CLASSES_ROOT, sKeyName & "\shell\" & sCommandName & "\command", , sCommandString)
    CreateExplorerShellCommand = bResult
End Function

 
'returns an extension's associated windows program
'Usage : x = GetAssociatedExecutable("vbp")
Public Function GetAssociatedExecutable(ByVal _
    Extension As String) As String

    Dim Path As String
    Dim filename As String
    Dim nRet As Long
    Const MAX_PATH As Long = 260
    
    'Create a tempfile
    Path = String$(MAX_PATH, 0)
    
    If GetTempPath(MAX_PATH, Path) Then
        filename = String$(MAX_PATH, 0)
    
        If GetTempFileName(Path, "~", 0, filename) Then
            filename = Left$(filename, _
                InStr(filename, vbNullChar) - 1)
        
            'Rename it to use supplied extension
            Name filename As Left$(filename, _
                InStr(filename, ".")) & Extension
                filename = Left$(filename, _
                InStr(filename, ".")) & Extension
        
            'Get name of associated EXE
            Path = String$(MAX_PATH, 0)
        
            Call FindExecutable(filename, _
                vbNullString, Path)
            GetAssociatedExecutable = Left$( _
                Path, InStr(Path, vbNullChar) - 1)
        
            'Clean up
            Kill filename
        
        End If
    
    End If
                                                                                                                                                                                                                                                               
End Function
'Purpose     :  Creates a key or sets an existing keys value in the registry
'Inputs      :  lRootKey                    A constant specifying which part of the registry to
'                                           write to, eg. HKEY_CLASSES_ROOT
'               sRegPath                    The path to write the value of the key to.
'               sValue                      The value of the key.
'Outputs     :
'Author      :  Andrew Baker
'Date        :  30/01/2001 11:53
'Notes       :  Used by CreateFileAssociation
'Revisions   :

Public Function zRegistryCreateKey(lRootKey As Long, sRegPath As String, Optional sValueName As String = "", Optional sValue As String = "") As Boolean
    Dim lhwnKey As Long
    Dim lRetVal As Long
    Const REG_SZ = 1
    
    On Error GoTo ErrFailed
    
    lRetVal = RegCreateKey(lRootKey, sRegPath, lhwnKey)
    If lRetVal = 0 Then
        'Successfully created/opened the key
        'Write value
        lRetVal = RegSetValueEx(lhwnKey, sValueName, 0, REG_SZ, ByVal sValue, Len(sValue))
        'Close key
        lRetVal = RegCloseKey(lhwnKey)
    End If
    zRegistryCreateKey = (lRetVal = 0)
    Exit Function

ErrFailed:
    zRegistryCreateKey = False
End Function
Function GetRegEntry(strKey As String, strSubKeys As String, strValName As String) As String
   'usage example: HardwareInfo = GetRegEntry("HKEY_LOCAL_MACHINE", "HARDWARE\DESCRIPTION\System\CentralProcessor\0", "Identifier")
   On Error GoTo GetRegEntry_Err
   '
   Dim lngType As Long, lngResult As Long, lngKey As Long
   Dim lngHandle As Long, lngcbData As Long
   Dim strRet As String
   Dim lngRet As Long
   '
   ' Take the human-readable key class and convert it to a WIN95 reserved
   ' numeric constant
   Select Case strKey
      Case "HKEY_CLASSES_ROOT"
         lngKey = &H80000000
      Case "HKEY_CURRENT_CONFIG"
         lngKey = &H80000005
      Case "HKEY_CURRENT_USER"
         lngKey = &H80000001
      Case "HKEY_DYN_DATA"
         lngKey = &H80000006
      Case "HKEY_LOCAL_MACHINE"
         lngKey = &H80000002
      Case "HKEY_PERFORMANCE_DATA"
         lngKey = &H80000004
      Case "HKEY_USERS"
         lngKey = &H80000003
      Case Else
         Exit Function
   End Select
   ' Open Key
   If Not ERROR_SUCCESS = RegOpenKeyEx(lngKey, strSubKeys, 0&, KEY_READ, lngHandle) Then
      Exit Function
   End If
   ' Get type of data in key value (lngtype)
   lngResult = RegQueryValueEx(lngHandle, strValName, 0&, lngType, ByVal strRet, lngcbData)
Select Case lngType
    Case 1    'if key data is string
        strRet = Space$(lngcbData)
        lngResult = RegQueryValueEx(lngHandle, strValName, 0&, lngType, ByVal strRet, lngcbData)

        If Not ERROR_SUCCESS = RegCloseKey(lngHandle) Then lngType = -1&
   
        GetRegEntry = strRet
    Case 3   'if key data is binary
        lngRet = 0&
        lngResult = RegQueryValueEx(lngHandle, strValName, 0&, 0&, lngRet, lngcbData)
        GetRegEntry = Str(lngRet)
End Select
   '
GetRegEntry_Exit:
   On Error GoTo 0
   Exit Function
   '
GetRegEntry_Err:
   lngType = -1&
   MsgBox Err & ">  " & error$, 16, "Common/GetRegEntry"
   Resume GetRegEntry_Exit
   '
End Function
Public Function GetString(Hkey As Long, strPath As String, strValue As String)
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    r = RegOpenKey(Hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

Public Sub SaveString(Hkey As Long, strPath As String, strValue As String, strData As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    r = RegCloseKey(keyhand)
End Sub

