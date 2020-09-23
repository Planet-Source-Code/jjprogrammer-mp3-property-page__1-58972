VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDropFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_nFiles As Long
Private m_sDropFiles() As String

Public Sub GetDropFiles(pDataObj As IDataObject, ByVal sExtension As String)

    Dim fmtEtc As FORMATETC
    Dim pmedium As STGMEDIUM
    Dim i As Long
    Dim lresult As Long
    Dim sTemp As String
    Dim lIndex As Long
    
    With fmtEtc
        .cfFormat = CF_HDROP
        .ptd = 0
        .dwAspect = DVASPECT_CONTENT
        .lIndex = -1
        .TYMED = TYMED_HGLOBAL
    End With
    
    pDataObj.GetData fmtEtc, pmedium
    
    m_nFiles = DragQueryFile(pmedium.pData, &HFFFFFFFF, vbNullString, 0)
        
    lIndex = 0
    For i = 0 To m_nFiles - 1
        sTemp = Space(255)
        lresult = DragQueryFile(pmedium.pData, i, sTemp, Len(sTemp))
        If (lresult > 0) Then
            sTemp = Left$(sTemp, lresult)
            If LCase(Right(sTemp, 4)) = sExtension Then
                ReDim Preserve m_sDropFiles(lIndex + 1)
                m_sDropFiles(lIndex) = sTemp
                lIndex = lIndex + 1
            End If
        End If
    Next i
    
    m_nFiles = lIndex
    
    ReleaseStgMedium pmedium
    
End Sub

Public Property Get Count() As Integer
    Count = m_nFiles
End Property

Public Property Get Files(nIndex As Integer) As String
    
    Files = ""
    
    If (m_nFiles) Then
        If (nIndex >= 0) And (nIndex < m_nFiles) Then
            Files = m_sDropFiles(nIndex)
        End If
    End If
        
End Property

Public Property Get SelectedFile() As String
    
    SelectedFile = ""
    
    If (m_nFiles) Then
        SelectedFile = m_sDropFiles(0)
    End If
    
End Property

