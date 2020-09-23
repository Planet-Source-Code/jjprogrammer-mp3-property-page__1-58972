Attribute VB_Name = "MP3TagModule"
Option Explicit

Public Type MP3TagInfo
  mTitle As String * 30
  mArtist As String * 30
  mAlbum As String * 30
  mYear As String * 4
  mComment As String * 30
  mGenre As String * 30
  mEncodedBy As String * 30
End Type

'ID3v1 Types
Private Type ID3v1Tag
    Identifier(2) As Byte
    Title(29) As Byte
    Artist(29) As Byte
    Album(29) As Byte
    SongYear(3) As Byte
    Comment(29) As Byte
    Genre As Byte
End Type

'ID3v2 Types
Private Type ID3v2Header
    Identifier(2) As Byte
    Version(1) As Byte
    Flags As Byte
    Size(3) As Byte
End Type

Private Type ID3v2ExtendedHeader
    Size(3) As Byte
End Type

Private Type ID3v2FrameHeader
    FrameID(3) As Byte
    Size(3) As Byte
    Flags(1) As Byte
End Type
Global MP3Filename As String

Public Type VBRinfo
  VBRrate As String
  VBRlength As String
End Type

Public Type MP3Info
  BITRATE As String
  channels As String
  COPYRIGHT As String
  CRC As String
  EMPHASIS As String
  freq As String
  LAYER As String
  length As String
  MPEG As String
  ORIGINAL As String
  Size As String
End Type
Public MyTag As MP3TagInfo
Public GenreName(1 To 148) As String
Public GenreNumber(1 To 148) As String
Global accMP3Info As MP3Info
Global mp3file As String
Private MP3Length As Long

Public Sub GetMP3Info(ByVal lpMP3File As String, ByRef lpMP3Info As MP3Info)
  Dim Buf As String * 4096
  Dim infoStr As String * 3
  Dim lpVBRinfo As VBRinfo
  Dim tmpByte As Byte
  Dim tmpNum As Byte
  Dim i As Integer
  Dim designator As Byte
  Dim baseFreq As Single
  Dim vbrBytes As Long
  
  Open lpMP3File For Binary As #1
    Get #1, 1, Buf
  Close #1
  
  For i = 1 To 4092
    If Asc(Mid(Buf, i, 1)) = &HFF Then
      tmpByte = Asc(Mid(Buf, i + 1, 1))
      If Between(tmpByte, &HF2, &HF7) Or Between(tmpByte, &HFA, &HFF) Then
        Exit For
      End If
    End If
  Next i
  If i = 4093 Then
   ' MsgBox "Not a MP3 file...", vbCritical, "Error..."
  Else
    infoStr = Mid(Buf, i + 1, 3)
    'Getting info from 2nd byte(MPEG,Layer type and CRC)
    tmpByte = Asc(Mid(infoStr, 1, 1))
    
    'Getting CRC info
    If ((tmpByte Mod 16) Mod 2) = 0 Then
      lpMP3Info.CRC = "Yes"
    Else
      lpMP3Info.CRC = "No"
    End If
    
    'Getting MPEG type info
    If Between(tmpByte, &HF2, &HF7) Then
      lpMP3Info.MPEG = "MPEG 2.0"
      designator = 1
    Else
      lpMP3Info.MPEG = "MPEG 1.0"
      designator = 2
    End If
    
    'Getting layer info
    If Between(tmpByte, &HF2, &HF3) Or Between(tmpByte, &HFA, &HFB) Then
      lpMP3Info.LAYER = "layer 3"
    Else
      If Between(tmpByte, &HF4, &HF5) Or Between(tmpByte, &HFC, &HFD) Then
        lpMP3Info.LAYER = "layer 2"
      Else
        lpMP3Info.LAYER = "layer 1"
      End If
    End If
    
    'Getting info from 3rd byte(Frequency, Bit-rate)
    tmpByte = Asc(Mid(infoStr, 2, 1))
    
    'Getting frequency info
    If Between(tmpByte Mod 16, &H0, &H3) Then
      baseFreq = 22.05
    Else
      If Between(tmpByte Mod 16, &H4, &H7) Then
        baseFreq = 24
      Else
        baseFreq = 16
      End If
    End If
    lpMP3Info.freq = baseFreq * designator * 1000 & " Hz"
    
    'Getting Bit-rate
    tmpNum = tmpByte \ 16 Mod 16
    If designator = 1 Then
      If tmpNum < &H8 Then
        lpMP3Info.BITRATE = tmpNum * 8
      Else
        lpMP3Info.BITRATE = 64 + (tmpNum - 8) * 16
      End If
    Else
      If tmpNum <= &H5 Then
        lpMP3Info.BITRATE = (tmpNum + 3) * 8
      Else
        If tmpNum <= &H9 Then
          lpMP3Info.BITRATE = 64 + (tmpNum - 5) * 16
        Else
          If tmpNum <= &HD Then
            lpMP3Info.BITRATE = 128 + (tmpNum - 9) * 32
          Else
            lpMP3Info.BITRATE = 320
          End If
        End If
      End If
    End If
    On Error Resume Next
    MP3Length = FileLen(lpMP3File) \ (Val(lpMP3Info.BITRATE) / 8) \ 1000
    If Mid(Buf, i + 36, 4) = "Xing" Then
      vbrBytes = Asc(Mid(Buf, i + 45, 1)) * &H10000
      vbrBytes = vbrBytes + (Asc(Mid(Buf, i + 46, 1)) * &H100&)
      vbrBytes = vbrBytes + Asc(Mid(Buf, i + 47, 1))
      GetVBRrate lpMP3File, vbrBytes, lpVBRinfo
      lpMP3Info.BITRATE = lpVBRinfo.VBRrate
      lpMP3Info.length = lpVBRinfo.VBRlength
    Else
      lpMP3Info.BITRATE = lpMP3Info.BITRATE & " Kbps"
      lpMP3Info.length = MP3Length & " seconds"
    End If
    
    'Getting info from 4th byte(Original, Emphasis, Copyright, Channels)
    tmpByte = Asc(Mid(infoStr, 3, 1))
    tmpNum = tmpByte Mod 16
    
    
    'Getting Copyright bit
    If tmpNum \ 8 = 1 Then
      lpMP3Info.COPYRIGHT = " Yes"
      tmpNum = tmpNum - 8
    Else
      lpMP3Info.COPYRIGHT = " No"
    End If
    
    'Getting Original bit
    If (tmpNum \ 4) Mod 2 Then
      lpMP3Info.ORIGINAL = " Yes"
      tmpNum = tmpNum - 4
    Else
      lpMP3Info.ORIGINAL = " No"
    End If
    
    'Getting Emphasis bit
    Select Case tmpNum
      Case 0
        lpMP3Info.EMPHASIS = " None"
      Case 1
        lpMP3Info.EMPHASIS = " 50/15 microsec"
      Case 2
        lpMP3Info.EMPHASIS = " invalid"
      Case 3
        lpMP3Info.EMPHASIS = " CITT j. 17"
    End Select
    
    'Getting channel info
    tmpNum = (tmpByte \ 16) \ 4
    Select Case tmpNum
      Case 0
        lpMP3Info.channels = " Stereo"
      Case 1
        lpMP3Info.channels = " Joint Stereo"
      Case 2
        lpMP3Info.channels = " 2 Channel"
      Case 3
        lpMP3Info.channels = " Mono"
    End Select
  End If
  lpMP3Info.Size = FileLen(lpMP3File) & " bytes"
End Sub

Private Sub GetVBRrate(ByVal lpMP3File As String, ByVal byteRead As Long, ByRef lpVBRinfo As VBRinfo)
  Dim i As Long
  Dim ok As Boolean

  i = 0
  byteRead = byteRead - &H39
  Do
    If byteRead > 0 Then
      i = i + 1
      byteRead = byteRead - 38 - Deljivo(i)
    Else
      ok = True
    End If
  Loop Until ok
  lpVBRinfo.VBRlength = Trim(str(i)) & " seconds"
  lpVBRinfo.VBRrate = Trim(str(Int(8 * FileLen(lpMP3File) / (1000 * i)))) & " Kbit (VBR)"
End Sub

Private Function Deljivo(ByVal Num As Long) As Byte
  If Num Mod 3 = 0 Then
    Deljivo = 1
  Else
    Deljivo = 0
  End If
End Function

Public Function Between(ByVal accNum As Byte, ByVal accDown As Byte, ByVal accUp As Byte) As Boolean
  If accNum >= accDown And accNum <= accUp Then
    Between = True
  Else
    Between = False
  End If
End Function



Public Function ReadID3v1(ByVal strFile As String, ByRef OutTag As MP3TagInfo) As Boolean
    Dim FileNo As Integer, fp As Long, i As Integer
    Dim RdTag As ID3v1Tag
    
    On Local Error GoTo Failed
    
    FileNo = FreeFile
    Open strFile For Binary As #FileNo
        fp = LOF(FileNo) - 127
        If fp > 0 Then
            Get #FileNo, fp, RdTag
            If GetStringValue(RdTag.Identifier, 3, 0) = "TAG" Then
                'An ID3v1 tag is present.
                OutTag.mTitle = Trim$(GetStringValue(RdTag.Title, 30, 0))
                OutTag.mArtist = Trim$(GetStringValue(RdTag.Artist, 30, 0))
                OutTag.mAlbum = Trim$(GetStringValue(RdTag.Album, 30, 0))
                OutTag.mYear = Trim$(GetStringValue(RdTag.SongYear, 4, 0))
                OutTag.mComment = Trim$(GetStringValue(RdTag.Comment, 30, 0))
                For i = 1 To 148
                    If GenreNumber(i) = RdTag.Genre Then Exit For
                Next i
                If i < 149 Then
                  OutTag.mGenre = GenreName(i)
                End If
                        ReadID3v1 = True
                    End If
                End If
    Close #FileNo
    Exit Function
    
Failed:
    Close #FileNo
    ReadID3v1 = False
End Function

Public Function WriteID3v1(ByVal mp3file As String, ByRef OutTag As MP3TagInfo) As Boolean
    Dim MP3Size As Long
    Dim locArtist As String * 30
    Dim locTitle As String * 30
    Dim locAlbum As String * 30
    Dim locYear As String * 4
    Dim locComment As String * 30
    Dim locGenre As String * 1
    Dim i As Integer

    On Local Error GoTo Failed
    
    MP3Size = FileLen(mp3file)
  With OutTag
        locTitle = .mTitle
        locArtist = .mArtist
        locAlbum = .mAlbum
        locYear = .mYear
        locComment = .mComment
        For i = 1 To 148
          If .mGenre = "" Then Exit For
          If Trim(.mGenre) = GenreName(i) Then GoTo TagIt
        Next i
        locGenre = Chr$(255)
TagIt:
        If i < 149 Then locGenre = Chr$(GenreNumber(i)) Else locGenre = Chr$(255)
  End With
  Open mp3file For Binary As #1
    Put #1, MP3Size + 1 - 128, "TAG" & locTitle & locArtist & locAlbum & locYear & locComment & locGenre
  Close #1

    WriteID3v1 = True
    Exit Function
    
Failed:
    Close #1
    WriteID3v1 = False
End Function

Public Function ReadID3v2(ByVal strFile As String, ByRef OutTag As MP3TagInfo) As Boolean
    Dim FileNo As Integer, fp As Long
    Dim RdHeader As ID3v2Header, RdExtHeader As ID3v2ExtendedHeader, RdFrameHeader As ID3v2FrameHeader
    Dim FrameID As String, FrameSize As Long, TextEncoding As Byte, RdData() As Byte, RdString As String
    Dim bGotArtist As Boolean, bGotTitle As Boolean, bGotAlbum As Boolean, bGotGenre As Boolean, bGotEnc As Boolean
    
    On Local Error GoTo Failed
    
    'Reads the ID3v2 tag of an mp3 file, if there is one.
    FileNo = FreeFile
    fp = 1
    Open strFile For Binary As #FileNo
        'Read the header.
        Get #FileNo, fp, RdHeader
        
        If GetStringValue(RdHeader.Identifier, 3, 0) = "ID3" Then
            fp = Loc(FileNo) + 1
            
            'An ID3v2 tag is present.
            If GetBit(6, RdHeader.Flags) Then
                'There is an extended header present. Just read its size to jump over it.
                Get #FileNo, , RdExtHeader
                fp = fp + GetLongValue(RdExtHeader.Size)
            End If
            
            Do
                Get #FileNo, fp, RdFrameHeader
                FrameID = GetStringValue(RdFrameHeader.FrameID, 4, 0)
                FrameSize = GetLongValue(RdFrameHeader.Size)
                If Not FrameSize < 2 Then
                    If FrameID = "TPE1" Or FrameID = "TIT2" Or FrameID = "TALB" Or FrameID = "TCON" Or FrameID = "TENC" Then
                        Get #FileNo, , TextEncoding
                        ReDim RdData(FrameSize - 2)
                        Get #FileNo, , RdData
                        RdString = GetStringValue(RdData, UBound(RdData) + 1, TextEncoding)
                        Select Case FrameID
                        Case "TPE1"
                            'Artist frame.
                            OutTag.mArtist = RdString
                            bGotArtist = True
                        Case "TIT2"
                            'Title frame.
                            OutTag.mTitle = RdString
                            bGotTitle = True
                        Case "TALB"
                            'Album frame.
                            OutTag.mAlbum = RdString
                            bGotAlbum = True
                        Case "TCON"
                            'Genre
                            OutTag.mGenre = RdString
                            bGotGenre = True
                        Case "TENC"
                            'Encoded By
                            OutTag.mEncodedBy = RdString
                            bGotGenre = True
                        End Select
                    End If
                End If
                'Seek to the next frame. The value + 10 is the frame header itself.
                fp = fp + 10 + FrameSize
            Loop While Not FrameSize = 0 And Not fp > 10 + GetLongValue(RdHeader.Size)
            If bGotArtist Or bGotTitle Or bGotAlbum Or bGotGenre Or bGotEnc Then ReadID3v2 = True
        End If
    Close #FileNo
    Exit Function
    
Failed:
    Close #FileNo
    ReadID3v2 = False
End Function

Public Function WriteID3v2(ByVal strFile As String, ByRef OutTag As MP3TagInfo) As Boolean
    Dim FileNo As Integer, fp As Long
    Dim AudioData() As Byte, AudioSize As Long, TagSize As Long
    Dim Header As ID3v2Header, WrHeader As ID3v2Header
    
    On Local Error GoTo Failed
    
    TagSize = Len(OutTag.mArtist) + Len(OutTag.mTitle) + Len(OutTag.mAlbum) + Len(OutTag.mGenre) + Len(OutTag.mEncodedBy)
    If Not Len(OutTag.mArtist) = 0 Then TagSize = TagSize + 11
    If Not Len(OutTag.mTitle) = 0 Then TagSize = TagSize + 11
    If Not Len(OutTag.mAlbum) = 0 Then TagSize = TagSize + 11
    If Not Len(OutTag.mGenre) = 0 Then TagSize = TagSize + 11
    If Not Len(OutTag.mEncodedBy) = 0 Then TagSize = TagSize + 11
    
    'Writes the ID3v2 tag of an mp3 file.
    FileNo = FreeFile
    fp = 1
    Open strFile For Binary As #FileNo
        AudioSize = LOF(FileNo)
        'Check for an existing header.
        Get #FileNo, fp, Header
        If GetStringValue(Header.Identifier, 3, 0) = "ID3" Then
            AudioSize = AudioSize - GetLongValue(Header.Size)
        End If
        'Save the existing audio data.
        ReDim AudioData(AudioSize - 1)
        Get #FileNo, LOF(FileNo) - AudioSize + 1, AudioData
    Close #FileNo
    Kill strFile
    Open strFile For Binary As #FileNo
        'Create the ID3 tag.
        '1) Create the header.
        SetStringValue WrHeader.Identifier, "ID3", 3
        WrHeader.Version(0) = 3
        SetLongValue WrHeader.Size, TagSize
        Put #FileNo, , WrHeader
        '2) Create the frames.
        WriteFrame FileNo, "TPE1", OutTag.mArtist
        WriteFrame FileNo, "TIT2", OutTag.mTitle
        WriteFrame FileNo, "TALB", OutTag.mAlbum
        WriteFrame FileNo, "TCON", OutTag.mGenre
        WriteFrame FileNo, "TENC", OutTag.mEncodedBy
        '3) Append the audio data.
        Put #FileNo, , AudioData
    Close #FileNo
    
    WriteID3v2 = True
    Exit Function
    
Failed:
    Close #FileNo
    WriteID3v2 = False
End Function

Private Sub WriteFrame(ByVal FileNo As Integer, ByVal strFrameHeader As String, ByVal strFrameData As String)
    Dim FrameHeader As ID3v2FrameHeader, EncData As Byte, FrameData() As Byte
    
    If Not Len(strFrameData) = 0 Then
        SetStringValue FrameHeader.FrameID, strFrameHeader, 4
        SetLongValue FrameHeader.Size, Len(strFrameData) + 1
        Put #FileNo, , FrameHeader
        ReDim FrameData(Len(strFrameData) - 1)
        SetStringValue FrameData, strFrameData, Len(strFrameData)
        Put #FileNo, , EncData
        Put #FileNo, , FrameData
    End If
End Sub

'Synchsafe integers are integers that keep its highest bit (bit 7) zeroed, making seven bits
'out of eight available. Thus a 32 bit synchsafe integer can store 28 bits of information.
Private Function GetLongValue(ByRef SyncsafeInt() As Byte) As Long
    Dim i As Integer, j As Integer, BitNr As Integer
    
    For i = 3 To 0 Step -1
        'Loop through the 4 bytes.
        For j = 0 To 6
            'Loop through the 7 significant bits per byte.
            If GetBit(j, SyncsafeInt(i)) Then
                GetLongValue = GetLongValue + 2 ^ BitNr
            End If
            BitNr = BitNr + 1
        Next j
    Next i
End Function

Private Sub SetLongValue(ByRef SyncsafeInt() As Byte, ByVal Value As Long)
    Dim i As Integer, ByteNr As Integer, BitNr As Integer
    
    ByteNr = 3
    For i = 0 To 27
        'Loop through the 28 bits of an synchsafe integer.
        If Value And 2 ^ i Then
            'This bit is set.
            SetBit BitNr, SyncsafeInt(ByteNr), True
        End If
        BitNr = BitNr + 1
        If BitNr Mod 7 = 0 Then
            'The next byte begins.
            ByteNr = ByteNr - 1
            BitNr = 0
        End If
    Next i
End Sub

Private Function GetStringValue(ByRef StringData() As Byte, ByVal StringLength As Integer, ByVal EncodingFormat As Byte) As String
    Dim i As Integer
    
    For i = 0 To StringLength - 1
        If EncodingFormat = 0 Or EncodingFormat = 3 Then
            'Clear text, null terminated.
            If StringData(i) = 0 Then Exit Function
            GetStringValue = GetStringValue & Chr$(StringData(i))
        ElseIf EncodingFormat = 1 Then
            'UNICODE text with BOM, double-null terminated.
            If i >= 2 And i Mod 2 = 0 Then
                If StringData(i) = 0 Then Exit Function
                GetStringValue = GetStringValue & Chr$(StringData(i))
            End If
        ElseIf EncodingFormat = 2 Then
            'UNICODE text without BOM, double-null terminated.
            If i Mod 2 = 0 Then
                If StringData(i) = 0 Then Exit Function
                GetStringValue = GetStringValue & Chr$(StringData(i))
            End If
        End If
        If Not EncodingFormat = 1 Or i >= 2 Then
        End If
    Next i
End Function

Private Sub SetStringValue(ByRef StringData() As Byte, ByVal Value As String, ByVal StringLength As Integer)
    Dim i As Integer
    
    For i = 0 To StringLength - 1
        StringData(i) = Asc(Mid$(Value, i + 1, 1))
    Next i
End Sub

'Bit Nr. 0 is the last bit, bit 7 the first bit.
Private Sub SetBit(ByVal BitNr As Integer, ByRef SrcData As Byte, ByVal BitState As Boolean)
    Dim Pattern As Byte
    
    If BitState Then
        'set a bit to 1
        Pattern = 2 ^ BitNr
        SrcData = SrcData Or Pattern
    Else
        'set a bit to 0
        Pattern = 255 - 2 ^ BitNr
        SrcData = SrcData And Pattern
    End If
End Sub

Private Function GetBit(ByVal BitNr As Byte, ByVal SrcData As Byte) As Boolean
    Dim Pattern As Byte
    
    Pattern = 2 ^ BitNr
    If SrcData And Pattern Then GetBit = True
End Function

Public Sub GetGenreData()
Dim genrestring$, Genre As Variant, x As Integer, t As Integer
genrestring$ = "123,A Cappella,34,Acid,74,Acid Jazz,73,Acid Punk,99,Acoustic,20,Alternative Rock,40,Alternative,26,Ambient,145,Anime,90,Avantgarde,116,Ballad,41,Bass,135,Beat,85,Bebob,96,Big Band,138,Black Metal,89,Bluegrass,0,Blues,107,Booty Bass,132,BritPop,65,Cabaret,88,Celtic,104,Chamber Music,102,Chanson,97,Chorus,136,Christian Gangsta Rap,61,Christian Rap,141,Christian Rock,32,Classical,1,Classic Rock,112,Club,128,Club-House,57,Comedy,140,Contemporary Christian,2,Country,139,Crossover,58,Cult,3,Dance,125,Dance Hall,50,Darkwave,22,Death Metal,4,Disco,55,Dream,127,Drum & Bass,122,Drum Solo,120,Duet,98,Easy Listening,52,Electronic,48,Ethnic,54,Eurodance,124,Euro-House,25,Euro-Techno,"
genrestring$ = genrestring$ & "84,Fast-Fusion,80,Folk,115,Folklore,81,Folk/Rock,119,Freestyle,5,Funk,30,Fusion,36,Game,59,Gangsta Rap,126,Goa,38,Gospel,49,Gothic,91,Gothic Rock,6,Grunge,129,Hardcore,79,Hard Rock, 137,Heavy Metal,7,Hip-Hop,35,House,100,Humour,131,Indie,19,Industrial,33,Instrumental,46,Instrumental Pop,47,Instrumental Rock,8,Jazz,29,Jazz+Funk, 146,JPop,63,Jungle,86,Latin,71,Lo-Fi,45,Meditative,142,Merengue,9,Metal,77,Musical,82,National Folk,64,Native American,133,Negerpunk,10,New Age,66,New Wave,39,Noise,11,Oldies,103,Opera,12,Other,75,Polka,134,Polsk Punk,13,Pop,53,Pop-Folk,62,Pop/Funk,109,Porn Groove,117,Power Ballad,23,Pranks,108,Primus,92,Progressive Rock,67,Psychedelic,"
genrestring$ = genrestring$ & "93,Psychedelic Rock,43,Punk,121,Punk Rock,15,Rap,68,Rave,14,R&B,16,Reggae,76,Retro,87,Revival,118,Rhythmic Soul,17,Rock,78,Rock & Roll,143,Salsa,114,Samba,110,Satire,69,Showtunes,21,Ska,111,Slow Jam, 95,Slow Rock,105,Sonata,42,Soul,37,Sound Clip,24,Soundtrack,56,Southern Rock,44,Space,101,Speech,83,Swing,94,Symphonic Rock,106,Symphony,147,Synthpop,113,Tango,18,Techno,51,Techno-Industrial,130,Terror,144,Thrash Metal,60,Top 40,70,Trailer,31,Trance,72,Tribal,27,Trip-Hop,28,Vocal,"
Genre = readin(genrestring$, -1, ",")
For x = 1 To UBound(Genre) + 1 Step 2
   t = t + 1
   GenreNumber(t) = Genre(x - 1)
   GenreName(t) = Genre(x)
Next x
End Sub
Function readin(ByVal Sourcestring As String, entry As Integer, Optional Delimiter As String = ";") As Variant
'Reads delimited data from Sourcestring
'syntax: value = Readin(a$,2) - reads 2nd entry in data string a$
'if Entry is < 0 then all data is returned in array
'if Entry is 0 then next data value is read
Static x As Integer
Dim item As String, temp As String, t As Integer, z As Integer
Dim RetArray As Variant
If entry < 0 Then
  RetArray = Empty
  If StrComp(Left$(Sourcestring, Len(Delimiter)), Delimiter, vbBinaryCompare) = 0 Then  'strip leading delimiter
       Sourcestring = Mid$(Sourcestring, Len(Delimiter) + 1, Len(Sourcestring))
       
    End If
    Do                                              ' loop to check for trailing delimiter(s)
     If StrComp(Right$(Sourcestring, Len(Delimiter)), Delimiter, vbBinaryCompare) = 0 Then  ' does the string have trailing delimiter?
        Sourcestring = Left$(Sourcestring, Len(Sourcestring) - Len(Delimiter))   ' strip trailing delimiter
        
     Else: Exit Do
     End If
     Loop
  t = 0
  z = 1
  Do              'get number of entries in t
    x = InStr(z, Sourcestring, Delimiter)
    If x = 0 Then Exit Do
    t = t + 1
    z = x + 1
  Loop
  ReDim RetArray(t) 'dim array to t
  t = 0
  z = 1
  
getentry:
  temp = ""
  Do                                'Extract entries
        x = x + 1
        item = Mid$(Sourcestring, x, 1)
        If item = Delimiter Or item = "" Then Exit Do
        temp = temp + item
        If x = Len(Sourcestring) Then x = 0: Exit Do
  Loop
    RetArray(t) = temp
    x = InStr(z, Sourcestring, Delimiter)
    
    If x = 0 Then GoTo leave
    t = t + 1
    z = x + 1
    GoTo getentry
leave:
    readin = RetArray
Exit Function
End If
If entry = 1 Then x = 0   'if entry = 1 then x is reset to 0
t = 0                     'if entry is 0 then x retains current value
z = 1
If entry > 1 Then          'if entry is 1 then skip following loop
  Do Until t = entry - 1   'Skip all entries before specified entry
    x = InStr(z, Sourcestring, Delimiter)
    t = t + 1
    z = x + 1
  Loop
End If
Start:
t = t + 1
Do                                'Extract specified entry
  x = x + 1
  item = Mid$(Sourcestring, x, 1)
  If item = Delimiter Or item = "" Then Exit Do
  temp = temp + item
  If x = Len(Sourcestring) Then x = 0: Exit Do
Loop
If entry > 0 And t <> entry Then
  temp = "": GoTo Start
Else: readin = temp
End If
End Function
Private Function ReplaceBadString(ByVal strData As String) As String
    Dim tmpstr As String
    tmpstr = strData
    'Replace invalid signs.
    tmpstr = Replace(tmpstr, "~", "_", , , vbTextCompare)
    tmpstr = Replace(tmpstr, "´", "'", , , vbTextCompare)
    tmpstr = Replace(tmpstr, "`", "'", , , vbTextCompare)
    tmpstr = Replace(tmpstr, "{", "(", , , vbTextCompare)
    tmpstr = Replace(tmpstr, "[", "(", , , vbTextCompare)
    tmpstr = Replace(tmpstr, "]", ")", , , vbTextCompare)
    tmpstr = Replace(tmpstr, "}", ")", , , vbTextCompare)
    tmpstr = Replace(tmpstr, "?", "¿", , , vbTextCompare)
    'Cut out invalid signs.
    tmpstr = Replace(tmpstr, "/", "", , , vbTextCompare)
    tmpstr = Replace(tmpstr, "\", "", , , vbTextCompare)
    tmpstr = Replace(tmpstr, ":", "", , , vbTextCompare)
    tmpstr = Replace(tmpstr, "*", "", , , vbTextCompare)
    tmpstr = Replace(tmpstr, """", "", , , vbTextCompare)
    tmpstr = Replace(tmpstr, "<", "", , , vbTextCompare)
    tmpstr = Replace(tmpstr, ">", "", , , vbTextCompare)
    tmpstr = Replace(tmpstr, "|", "", , , vbTextCompare)
     ReplaceBadString = tmpstr
End Function
 Private Function Replace(ByVal StrOriginal As String, ByVal StrFind As String, ByVal StrReplace As String, Optional ByVal intOPMode As Integer, Optional Updated As Integer, Optional method As Integer = vbTextCompare) As String

  '*******************************************************************************
  '
  ' DESCRIPTION
  '     Replace a string or specific character(s) within a string. This routine
  '     can also be used to strip characters.
  '
  ' ARGUMENTS
  '     StrOriginal    = String to work on.
  '
  '     StrFind     defines string to search for.
  
  '     StrReplace     = New character (or string) to substitute.
  '
  '     intOPMode  = Sets operation by defining the "replace" mode and "compare"
  '                  mode. Valid parameters are:
  '
  '                  BinaryCompare (Case sensitive. Default if not specified.)
  '                  TextCompare (Not case sensitive)
  '                  DataBaseCompare (Microsoft Access data compare)
  '
  '    Updated = Optional. Returns positive if string was modified. Value is number
  '                  of replacements made.
  '
  ' RETURNS
  '     Returns new string.
  '
  ' DEPENDENCIES
  '     None
  '
  ' REMARKS
  '     To strip a string of character(s), set StrReplace to vbNullString or "".
  '
  '*******************************************************************************
  
  On Local Error Resume Next
  
  Dim intOldLen As Integer
  Dim intNewLen As Integer
  Dim intSPos As Long
  Dim intN As Integer
  
  intNewLen = Len(StrReplace)
  intOldLen = Len(StrFind)
    
  intSPos = 1
 Updated = 0
    
  
    Do
      intSPos = InStr(intSPos, StrOriginal, StrFind, intOPMode)
      If intSPos Then
        StrOriginal = Left(StrOriginal, intSPos - 1) & StrReplace & Mid(StrOriginal, intSPos + intOldLen)
        intSPos = intSPos + intNewLen
       Updated = Updated + 1
      End If
    Loop While intSPos
  
  Replace = StrOriginal
  
End Function
Public Function GetTagInfo(MP3Filename As String, TmpTag As MP3TagInfo) As Boolean
Dim MP3Size As Long
  
  If Not FileExists(MP3Filename) Then Exit Function
  MP3Size = FileLen(MP3Filename)
  If ReadID3v1(MP3Filename, TmpTag) Then GetTagInfo = True
  If ReadID3v2(MP3Filename, TmpTag) Then GetTagInfo = True

End Function
Public Function FileExists(ByVal filename As String) As Boolean
    If Not filename > "" Then
        FileExists = False
        Exit Function
    End If
    On Error Resume Next
    FileExists = Dir$(filename) <> ""
End Function
 Public Function GetEncoder(ByVal filename As String) As String
Dim output$, encoder As String
On Error Resume Next
output$ = ShellExecuteCapture("EncSpotDOS " & Chr$(34) & filename & Chr$(34))
encoder = Mid$(output$, InStr(output$, "Encoder") + 21, 30)
GetEncoder = Mid$(encoder, 1, InStr(1, encoder, vbCr))
End Function

