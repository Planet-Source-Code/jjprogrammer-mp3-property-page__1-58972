Attribute VB_Name = "PropSheet"
Option Explicit

'Window Message Constants
Public Const WM_COMMAND = &H111
Public Const WM_DESTROY = &H2
Public Const WM_INITDIALOG = &H110
Public Const WM_NOTIFY = &H4E

'Empty Dialog template
Public Const IDD_MP3PROPDLG = 100
'Property Sheet object reference
Public m_pPropSheet As clsPropSheet

Public Function PropSheetCallbackProc(ByVal hwnd As hwnd, ByVal uMsg As UINT, ByVal ppsp As LPPROPSHEETPAGE) As Long

    'Not implemented. This is a skeleton
    
    Dim psp As PROPSHEETPAGE
    CopyMemory psp, ByVal ppsp, Len(psp)
    
    'Get reference to object. No AddRef!!!!
    CopyMemory m_pPropSheet, psp.lparam, 4
    
    Select Case uMsg
        
        Case PSPCB_CREATE
            'Return non-zero to create page. 0 prevents it.

        Case PSPCB_RELEASE:
            'Page is being destroyed. Return value is ignored.
            
    End Select
    
    
    PropSheetCallbackProc = 1
        
End Function

Public Function PropSheetDlgProc(ByVal hwndDlg As hwnd, ByVal uMsg As UINT, ByVal wParam As wParam, ByVal lparam As lparam) As BOOL

    Select Case uMsg
        
        Case WM_INITDIALOG
            Load frmTagProperties
            
            ' Move the form to the
            ' property page
            SetParentAPI frmTagProperties.hwnd, hwndDlg

            ' Pass the page handle to
            ' the form
            m_lSheet = hwndDlg

            ' Pass the dialog handle to
            ' the from
            m_lPropSheet = GetParent(hwndDlg)

            ' Show the form
            SetWindowPos frmTagProperties.hwnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_SHOWWINDOW
            
        Case WM_NOTIFY
            Notify hwndDlg, lparam
            
        Case WM_DESTROY
            'DO NOT DO THIS: Set m_pPropSheet = Nothing.
            CopyMemory m_pPropSheet, 0&, 4
    
    End Select
    
    PropSheetDlgProc = 0
    
End Function



Public Sub Notify(ByVal hwndDlg As hwnd, ByVal lparam As lparam)

    Dim nh As NMHDR
    
    CopyMemory nh, ByVal lparam, Len(nh)
    
    Select Case nh.code
        Case PSN_APPLY
            'OK and Apply
            
        Case PSN_QUERYCANCEL
            'Cancel has been clicked. Return 1 to prevent. 0 to allow.
            
        Case PSN_SETACTIVE
            'sent when property tab is selected for first time
            
        Case PSN_KILLACTIVE
            'sent when another property tab is selected
            
        Case PSN_RESET
            'Cancel has been allowed. About to be destroyed
        
    End Select
            
End Sub

