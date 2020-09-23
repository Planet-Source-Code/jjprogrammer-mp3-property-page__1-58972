VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPropSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IShellExtInit
Implements IShellPropSheetExt

Private m_clsDropFiles As clsDropFiles

Private Sub IShellExtInit_Initialize(ByVal pidlFolder As VBShellLib.LPCITEMIDLIST, ByVal pDataObj As VBShellLib.IDataObject, ByVal hKeyProgID As VBShellLib.Hkey)

    Set m_clsDropFiles = New clsDropFiles
    m_clsDropFiles.GetDropFiles pDataObj, ".mp3"
    
End Sub

Private Sub IShellPropSheetExt_AddPages(ByVal lpfnAddPage As VBShellLib.LPFNADDPROPSHEETPAGE, ByVal lparam As VBShellLib.lparam)

    Dim psp As PROPSHEETPAGE
    Dim sTitle As String
    Dim lAddPage As Long
    
    sTitle = StrConv("MP3 Tag Settings", vbFromUnicode)
    
    With psp
        .dwSize = Len(psp)
        .dwFlags = PSP_USECALLBACK Or PSP_USETITLE
        .hInstance = App.hInstance
        .lparam = ObjPtr(Me)
        .pfnDlgProc = GetAddress(AddressOf PropSheetDlgProc)
        .pfnCallback = GetAddress(AddressOf PropSheetCallbackProc)
        .pszTemplate = IDD_MP3PROPDLG 'Empty Dialog to be replaced
        .pszTitle = StrPtr(sTitle)
    End With
    
    'AddRef
    Dim pUnk As IUnknownVB
    Set pUnk = Me
    pUnk.AddRef
    
    Dim hPage As Long
    hPage = CreatePropertySheetPage(psp)
    
    
    If hPage Then
        lAddPage = AddPropertyPage(lpfnAddPage, hPage, lparam)
        
        If lAddPage = 0 Then
            DestroyPropertySheetPage hPage
        End If
    End If
    
    
End Sub

Private Sub IShellPropSheetExt_ReplacePage(ByVal uPageID As VBShellLib.UINT, ByVal lpfnReplaceWith As VBShellLib.LPFNADDPROPSHEETPAGE, ByVal lparam As VBShellLib.lparam)
    Err.Raise E_NOTIMPL
End Sub

Public Property Get SelectedFile() As String
    SelectedFile = ""
    If (m_clsDropFiles.Count) Then
        SelectedFile = m_clsDropFiles.SelectedFile
    End If
End Property
