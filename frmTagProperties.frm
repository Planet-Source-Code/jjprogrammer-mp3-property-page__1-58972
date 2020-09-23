VERSION 5.00
Begin VB.Form frmTagProperties 
   BorderStyle     =   0  'None
   ClientHeight    =   6120
   ClientLeft      =   4935
   ClientTop       =   1755
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   341
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   510
      Left            =   900
      TabIndex        =   24
      Top             =   2580
      Width           =   3285
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Default         =   -1  'True
         Height          =   375
         Left            =   15
         TabIndex        =   1
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Reset"
         Height          =   375
         Left            =   1065
         TabIndex        =   0
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove ID3"
         Height          =   375
         Left            =   2130
         TabIndex        =   25
         Top             =   30
         Width           =   1095
      End
   End
   Begin VB.Frame frmID3 
      Caption         =   "Tag info"
      Height          =   2055
      Left            =   510
      TabIndex        =   8
      Top             =   375
      Width           =   3975
      Begin VB.TextBox Comment 
         Height          =   285
         Left            =   840
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox Title 
         Height          =   285
         Left            =   840
         MaxLength       =   30
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox Artist 
         Height          =   285
         Left            =   840
         MaxLength       =   30
         TabIndex        =   3
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox Album 
         Height          =   285
         Left            =   840
         MaxLength       =   30
         TabIndex        =   4
         Top             =   945
         Width           =   3015
      End
      Begin VB.TextBox Year 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox Genre 
         Height          =   315
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   6
         Text            =   "Genre"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label labGenre 
         AutoSize        =   -1  'True
         Caption         =   "Genre"
         Height          =   195
         Left            =   1440
         TabIndex        =   14
         Top             =   1365
         Width           =   435
      End
      Begin VB.Label labYear 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Year"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1365
         Width           =   450
      End
      Begin VB.Label labComent 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Comment"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   1725
         Width           =   660
      End
      Begin VB.Label labTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Title"
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   300
         Width           =   420
      End
      Begin VB.Label labAlbum 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Album"
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   990
         Width           =   435
      End
      Begin VB.Label labArtist 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Artist"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   645
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add ID3"
      Height          =   375
      Left            =   2250
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2610
      Width           =   975
   End
   Begin VB.Frame frmNoTag 
      Caption         =   "Tag info"
      Height          =   2055
      Left            =   510
      TabIndex        =   20
      Top             =   375
      Width           =   3975
      Begin VB.Label labNoTag 
         AutoSize        =   -1  'True
         Caption         =   "This MP3 doesn't contain ID3 Tag"
         Height          =   195
         Left            =   360
         TabIndex        =   21
         Top             =   600
         Width           =   2430
      End
   End
   Begin VB.Frame frmMPEG 
      Caption         =   "MPEG info"
      Height          =   2475
      Left            =   885
      TabIndex        =   15
      Top             =   3225
      Width           =   3435
      Begin VB.Label LabEnc 
         Caption         =   "Encoded by:"
         Height          =   420
         Left            =   120
         TabIndex        =   32
         Top             =   2010
         Width           =   3270
      End
      Begin VB.Label labEmphasis 
         AutoSize        =   -1  'True
         Caption         =   "Emphasis:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1815
         Width           =   720
      End
      Begin VB.Label labOriginal 
         AutoSize        =   -1  'True
         Caption         =   "Original:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   1620
         Width           =   570
      End
      Begin VB.Label labCopy 
         AutoSize        =   -1  'True
         Caption         =   "Copyrighted:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1425
         Width           =   885
      End
      Begin VB.Label labBitRate 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   825
         Width           =   600
      End
      Begin VB.Label labLayer 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   615
         Width           =   720
      End
      Begin VB.Label labCRC 
         AutoSize        =   -1  'True
         Caption         =   "CRCs:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1215
         Width           =   450
      End
      Begin VB.Label labFreqChan 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label labLength 
         AutoSize        =   -1  'True
         Caption         =   "Length:"
         Height          =   195
         Left            =   105
         TabIndex        =   18
         Top             =   435
         Width           =   540
      End
      Begin VB.Label labSize 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   345
      End
   End
   Begin VB.Frame frmNoMPEG 
      Caption         =   "MPEG info"
      Height          =   2415
      Left            =   1590
      TabIndex        =   22
      Top             =   3210
      Width           =   2175
      Begin VB.Label labNoMPEG 
         AutoSize        =   -1  'True
         Caption         =   "Probably not a MP3 file..."
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1770
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   225
      Left            =   45
      TabIndex        =   33
      Top             =   105
      Width           =   5040
   End
End
Attribute VB_Name = "frmTagProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MP3Size As Long
Private Function GetEncoder(ByVal filename As String) As String
Dim output$, encoder As String
output$ = ShellExecuteCapture("EncSpotDOS " & Chr$(34) & filename & Chr$(34))
encoder = Mid$(output$, InStr(output$, "Encoder") + 21, 30)
GetEncoder = Mid$(encoder, 1, InStr(1, encoder, vbCr))
End Function

Public Function GetLongFileName(ByVal ShortFileName As String) As String

    Dim intPos As Integer
    Dim strLongFileName As String
    Dim strDirName As String
    
    'Format the filename for later processing
    ShortFileName = ShortFileName & "\"
    
    'Grab the position of the first real slash
    intPos = InStr(4, ShortFileName, "\")
    
    'Loop round all the directories and files
    'in ShortFileName, grabbing the full names
    'of everything within it.
    
    While intPos
    
        strDirName = Dir(Left(ShortFileName, intPos - 1), _
            vbNormal + vbHidden + vbSystem + vbDirectory)
        
        If strDirName = "" Then
            GetLongFileName = ""
            Exit Function
        End If
        
        strLongFileName = strLongFileName & "\" & strDirName
        intPos = InStr(intPos + 1, ShortFileName, "\")
        
    Wend

    'Return the completed long file name
    GetLongFileName = Left(ShortFileName, 2) & strLongFileName
  
End Function
Private Sub Album_GotFocus()
selectalltext Album
End Sub

Private Sub Artist_GotFocus()
selectalltext Artist
End Sub

Private Sub cmdAdd_Click()
  Dim emptyStr As String * 124
  
  frmID3.Visible = True
  frmButtons.Visible = True
  MP3Size = FileLen(mp3file)
  Open mp3file For Binary As #1
    Put #1, MP3Size, "Dejvi"
    Put #1, MP3Size, Chr$(0) & "TAG" & emptyStr & Chr$(255)
  Close #1
End Sub

Private Sub cmdCancel_Click()
Call Label1_Change
End Sub

Private Sub cmdRemove_Click()
  Dim tmpstr As String
  If ReadID3v1(mp3file, MyTag) = False Then GoTo V2
  MP3Size = FileLen(mp3file)
  tmpstr = Space(MP3Size - 128)
  Open mp3file For Binary As #1
    Get #1, 1, tmpstr
  Close #1
  Kill mp3file
  Open mp3file For Binary As #1
    Put #1, 1, tmpstr
  Close #1
V2:
  If ReadID3v2(mp3file, MyTag) = False Then Exit Sub
  With MyTag
    .mAlbum = Space$(30)
    .mArtist = Space$(30)
    .mComment = Space$(30)
    .mGenre = Chr$(0)
    .mTitle = Space$(30)
    .mYear = Space$(4)
  End With
  WriteID3v1 mp3file, MyTag
  WriteID3v2 mp3file, MyTag
  cmdCancel_Click
End Sub

Private Sub cmdSave_Click()
  MP3Size = FileLen(mp3file)
  With MyTag
    .mAlbum = Album
    .mArtist = Artist
    .mComment = Comment
    .mGenre = Genre
    .mTitle = Title
    .mYear = Year
  End With
  WriteID3v1 mp3file, MyTag
  WriteID3v2 mp3file, MyTag
End Sub


Public Sub GetMP3Inf()
  Dim Duration As Double, accMP3Info As MP3Info
  If Not FileExists(mp3file) Then Exit Sub
  GetMP3Info mp3file, accMP3Info
  labSize = "Size:                " & Format$(Val(accMP3Info.Size) / 1048576, "#####.0#") & " Mb"
  Duration = Val(accMP3Info.length)
  If Duration > 60 Then
                  Min = Format$(Duration \ 60, "###")
                  sec = Format$(Duration - (Min * 60), "0#")
              Else
                 Min = 0
                 sec = Format$(Duration, "0#")
              End If
              If sec = 60 Then Exit Sub
  labLength = "Duration:          " & Min & ":" & Format$(sec, "0#")
  labLayer = "Type:               " & accMP3Info.MPEG & " " & accMP3Info.LAYER
  labBitRate = "Bitrate:            " & accMP3Info.BITRATE
  labFreqChan = "Frequency:      " & accMP3Info.freq & " " & accMP3Info.channels
  labCRC = "CRC's:             " & accMP3Info.CRC
  labCopy = "Copyrighted:   " & accMP3Info.COPYRIGHT
  labEmphasis = "Emphasis:       " & accMP3Info.EMPHASIS
  labOriginal = "Original:          " & accMP3Info.ORIGINAL
  LabEnc = "Encoded by:     "
  If Trim(MyTag.mEncodedBy) > " " Then
    LabEnc = "Encoded by:    " & Trim(MyTag.mEncodedBy)
  Else
    MyTag.mEncodedBy = GetEncoder(mp3file)
    LabEnc = "Encoded by:    " & Trim(MyTag.mEncodedBy)
    WriteID3v2 mp3file, MyTag
  End If
End Sub
Public Sub GetTagInf()
  Dim i As Byte
  
  'Get the size of mp3 file(in bytes)
  If Not FileExists(mp3file) Then Exit Sub
  MP3Size = FileLen(mp3file)
  If GetTagInfo(mp3file, MyTag) Then
      frmID3.Visible = True
      frmButtons.Visible = True
      With MyTag
        Title = .mTitle
        Artist = .mArtist
        Album = .mAlbum
        Year = .mYear
        Comment = .mComment
      For i = 0 To 147
        If Genre.List(i) = Trim(.mGenre) Then Exit For
      Next i
      If i < 147 Then
        Genre.ListIndex = i
      Else
        Genre = .mGenre
      End If
      End With
  End If

End Sub

Private Sub Comment_GotFocus()
selectalltext Comment
End Sub

Private Sub Form_Load()
  GetGenreData
  For x = 1 To 148
    Genre.AddItem GenreName(x)
  Next x
  mp3file = m_pPropSheet.SelectedFile
  Label1.Caption = m_pPropSheet.SelectedFile 'label1_change event triggers MP3Info
End Sub

Private Sub Label1_Change()
GetTagInf
GetMP3Inf
End Sub
Sub selectalltext(txtbox As Control)
txtbox.SelStart = 0
txtbox.SelLength = 6500
End Sub

Private Sub Title_GotFocus()
selectalltext Title
End Sub

Private Sub Year_GotFocus()
selectalltext Year
End Sub
