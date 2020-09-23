Attribute VB_Name = "handler"
Option Explicit
' Window Styles
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000

Public Const E_NOTIMPL = &H80004001
Public Const IDCANCEL = 2
Public Const IDNO = 7
Public Const IDYES = 6
Public Const GMEM_MOVEABLE = &H2
Public Const MAX_PATH = 260
Public Const PAGE_EXECUTE_READWRITE = &H40&
Public Const S_FALSE = 1
Public Const S_OK = 0
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

Public Const WS_EX_CONTROLPARENT = &H10000

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Type DLGTEMPLATE
    Style As Long
    dwExtendedStyle As Long
    cdit As Integer
    x As Integer
    y As Integer
    cx As Integer
    cy As Integer
    Menu As Integer
    Class As String * 7
    Caption As Integer
End Type
Public g_tDlgTemplate As DLGTEMPLATE
Public m_lPropSheet As Long
Public m_lSheet As Long

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function AddPropertyPage Lib "propext.dll" (ByVal lpfn As Long, ByVal hPage As Long, ByVal lparam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CreatePropertySheetPage Lib "comctl32.dll" Alias "CreatePropertySheetPageA" (p As PROPSHEETPAGE) As Long
Public Declare Sub DestroyPropertySheetPage Lib "comctl32.dll" (ByVal hPage As Long)
Public Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Long, ByVal pUINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Public Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function OleRegEnumFormatEtc Lib "ole32.dll" (refclsid As GUID, ByVal dwDirection As DATADIR, lpEnumFormatEtc As IEnumFORMATETC) As Long
Public Declare Function ReleaseStgMedium Lib "ole32.dll" (pmedium As STGMEDIUM) As Long
Public Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Integer, ByVal nIDDlgItem As Integer, ByVal lpString As String) As Integer
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Public Declare Function WPPString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function SetParentAPI Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Boolean


Public Function LOWORD(ByVal lVal As Long) As Integer
    LOWORD = lVal And &HFFFF&
End Function

Public Function HIWORD(ByVal lVal As Long) As Integer
    
    HIWORD = 0
    If lVal Then
        HIWORD = lVal \ &H10000 And &HFFFF&
    End If
   
End Function

Public Function GetAddress(ByVal lpfn As Long)
    GetAddress = lpfn
End Function

Public Sub StrFromPtrW(pOLESTR As Long, strOut As String)
    Dim ByteArray(255) As Byte
    Dim intTemp As Integer
    Dim intCount As Integer
    Dim i As Integer
    
    intTemp = 1
    
    'Walk the string and retrieve the first byte of each WORD.
    While intTemp <> 0
        CopyMemory intTemp, ByVal pOLESTR + i, 2
        ByteArray(intCount) = intTemp
        intCount = intCount + 1
        i = i + 2
    Wend
    
    'Copy the byte array to our string.
    CopyMemory ByVal strOut, ByteArray(0), intCount
End Sub

Public Sub StrFromPtrA(pOLESTR As Long, strOut As String)
    Dim ByteArray(255) As Byte
    Dim intTemp As Integer
    Dim intCount As Integer
    Dim i As Integer
    
    intTemp = 1
    
    'Walk the string and retrieve the first byte of each WORD.
    While intTemp <> 0
        CopyMemory intTemp, ByVal pOLESTR + i, 1
        ByteArray(intCount) = intTemp
        intCount = intCount + 1
        i = i + 1
    Wend
    
    'Copy the byte array to our string.
    CopyMemory ByVal strOut, ByteArray(0), intCount
    
End Sub

Public Function SwapVtableEntry(pObj As Long, EntryNumber As Integer, ByVal lpfn As Long) As Long

    Dim lOldAddr As Long
    Dim lpVtableHead As Long
    Dim lpfnAddr As Long
    Dim lOldProtect As Long

    CopyMemory lpVtableHead, ByVal pObj, 4
    lpfnAddr = lpVtableHead + (EntryNumber - 1) * 4
    CopyMemory lOldAddr, ByVal lpfnAddr, 4

    Call VirtualProtect(lpfnAddr, 4, PAGE_EXECUTE_READWRITE, lOldProtect)
    CopyMemory ByVal lpfnAddr, lpfn, 4
    Call VirtualProtect(lpfnAddr, 4, lOldProtect, lOldProtect)

    SwapVtableEntry = lOldAddr

End Function

Public Function TrimNull(str As String) As String
    str = Trim(str)
    TrimNull = Left(str, Len(str) - 1)
End Function

Public Sub Log(sMsg As String)

    Dim hFile As Integer
    hFile = FreeFile

    Open App.Path & "/" & "raddata.log" For Append As #hFile

    Write #hFile, sMsg
    Close #hFile

End Sub
