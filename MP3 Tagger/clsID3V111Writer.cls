VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsID3V111Writer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ID3Info
    strHeader As String * 3
    strSongTitle As String * 30
    strArtist As String * 30
    strAlbum As String * 30
    strYear As String * 4
    strComment As String * 28
    bytUnused As Byte
    bytTrack As Byte
    bytGenre As Byte
End Type

Private fstrGenres() As Variant

Private fstrSongTitle As String
Private fstrArtist As String
Private fstrAlbum As String
Private fstrYear As String
Private fstrComment As String
Private fbytTrack As Byte
Private fbytGenre As Byte
Private fbolGenreSet As Boolean

Public Property Let SongTitle(ByVal strSongTitle As String)
   fstrSongTitle = strSongTitle
End Property

Public Property Get SongTitle() As String
    SongTitle = fstrSongTitle
End Property

Public Property Let Artist(ByVal strArtist As String)
    fstrArtist = strArtist
End Property

Public Property Get Artist() As String
    Artist = fstrArtist
End Property

Public Property Let Album(ByVal strAlbum As String)
    fstrAlbum = strAlbum
End Property

Public Property Get Album() As String
    Album = fstrAlbum
End Property

Public Property Let Year(ByVal strYear As String)
    fstrYear = strYear
End Property

Public Property Get Year() As String
    Year = fstrYear
End Property

Public Property Let Comment(ByVal strComment As String)
    fstrComment = strComment
End Property

Public Property Get Comment() As String
    Comment = fstrComment
End Property

Public Property Let Track(ByVal bytTrack As Byte)
    fbytTrack = bytTrack
End Property

Public Property Get Track() As Byte
    Track = fbytTrack
End Property

Public Property Let Genre(ByVal bytGenre As Byte)
    fbytGenre = bytGenre
    fbolGenreSet = True
End Property

Public Property Get Genre() As Byte
    Genre = fbytGenre
End Property

Public Sub RemoveID3V111Tag(ByVal strFileName As String)
    Dim intFile As Integer
    Dim intFile2 As Integer
    Dim lngTotalBytesToRead As Long
    Dim lngBuffer As Long
    Dim bytBytes() As Byte
    
    On Error GoTo ErrHan
    
    If FileExists(strFileName) = False Then Exit Sub
    If HasID3V111Tag(strFileName) = False Then Exit Sub
    
    intFile = FreeFile
    Open strFileName For Binary Access Read Shared As intFile
    intFile2 = FreeFile
    Open strFileName & ".mp3" For Binary Access Write Shared As intFile2
        lngTotalBytesToRead = LOF(intFile) - 128
        lngBuffer = 2097152
        ReDim bytBytes(1 To lngBuffer)
        
        Do Until EOF(intFile)
            If lngTotalBytesToRead < lngBuffer Then
                lngBuffer = lngTotalBytesToRead
                ReDim bytBytes(1 To lngBuffer)
            End If
            
            Get #intFile, , bytBytes
            Put #intFile2, , bytBytes
            lngTotalBytesToRead = lngTotalBytesToRead - lngBuffer
            If lngTotalBytesToRead <= 0 Then Exit Do
            DoEvents
        Loop
    Close intFile2
    Close intFile
    
    Kill strFileName
    Name strFileName & ".mp3" As strFileName
    
ErrHan:
    Close intFile
    Close intFile2
    If FileExists(strFileName & ".mp3") = True Then Kill strFileName & ".mp3"
End Sub

Public Sub WriteID3V111Tag(ByVal strFileName As String)
    Dim intFile As Integer
    Dim ID3Tag As ID3Info
    
    On Error GoTo ErrHan
    
    If FileExists(strFileName) = False Then Exit Sub
    If HasID3V111Tag(strFileName) = True Then RemoveID3V111Tag (strFileName)
    
    If fbolGenreSet = False Then
        Genre = 255
        fbolGenreSet = False
    End If
    
    With ID3Tag
        .strHeader = "TAG"
        .strSongTitle = SongTitle
        .strArtist = Artist
        .strAlbum = Album
        .strYear = Year
        .strComment = Comment
        .bytUnused = 0
        .bytTrack = Track
        .bytGenre = Genre
    End With
    
    intFile = FreeFile
    Open strFileName For Binary Access Write As intFile
        Put #intFile, LOF(intFile) + 1, ID3Tag
    Close intFile
    
    Exit Sub
    
ErrHan:
End Sub

Private Function HasID3V111Tag(ByVal strFileName As String) As Boolean
    Dim intFile As Integer
    Dim ID3 As ID3Info
    
    On Error GoTo ErrHan
    
    If FileExists(strFileName) = False Then Exit Function
    
    intFile = FreeFile
    Open strFileName For Binary Access Read Shared As intFile
        Get #intFile, LOF(intFile) - 127, ID3
    Close intFile
    
    With ID3
        If .strHeader = "TAG" Then HasID3V111Tag = True
    End With
    
    Exit Function
    
ErrHan:
End Function

Private Function FileExists(ByVal strFileName As String) As Boolean
    Dim intFile As Integer
    
    On Error GoTo ErrHan
    
    intFile = FreeFile
    Open strFileName For Input As intFile
    Close intFile
    
    FileExists = True
    Exit Function
    
ErrHan:
    FileExists = False
End Function

Public Function GetGenre(ByVal strGenre As String) As Byte
    Dim lngCounter As Long
    
    On Error GoTo ErrHan
    
    For lngCounter = 0 To UBound(fstrGenres)
        If LCase(strGenre) = LCase(fstrGenres(lngCounter)) Then
            GetGenre = lngCounter
            Exit Function
        End If
    Next
    
    If lngCounter = UBound(fstrGenres) + 1 Then GetGenre = 255
    
    Exit Function
    
ErrHan:
    GetGenre = 255
End Function

Private Sub Class_Initialize()
    fstrGenres = Array("Blues", "Classic Rock", "Country", "Dance", "Disco", "Funk", "Grunge", "Hip-Hop", _
                       "Jazz", "Metal", "New Age", "Oldies", "Other", "Pop", "R&B", "Rap", "Reggae", "Rock", _
                       "Techno", "Industrial", "Alternative", "Ska", "Death Metal", "Pranks", "Soundtrack", _
                       "Euro-Techno", "Ambient", "Trip-Hop", "Vocal", "Jazz+Funk", "Fusion", "Trance", _
                       "Classical", "Instrumental", "Acid", "House", "Game", "Sound Clip", "Gospel", "Noise", _
                       "AlternRock", "Bass", "Soul", "Punk", "Space", "Meditative", "Instrumental Pop", _
                       "Instrumental Rock", "Ethnic", "Gothic", "Darkwave", "Techno-Industrial", "Electronic", _
                       "Pop-Folk", "Eurodance", "Dream", "Southern Rock", "Comedy", "Cult", "Gangsta", _
                       "Top 40", "Christian Rap", "Pop/Funk", "Jungle", "Native American", "Cabaret", _
                       "New Wave", "Psychadelic", "Rave", "Showtunes", "Trailer", "Lo-Fi", "Tribal", _
                       "Acid Punk", "Acid Jazz", "Polka", "Retro", "Musical", "Rock & Roll", "Hard Rock", _
                       "Folk", "Folk/Rock", "National folk", "Swing", "Fast-fusion", "Bebob", "Latin", _
                       "Revival", "Celtic", "Bluegrass", "Avantgarde", "Gothic Rock", "Progressive Rock", _
                       "Psychedelic Rock", "Symphonic Rock", "Slow Rock", "Big Band", "Chorus", _
                       "Easy Listening", "Acoustic", "Humour", "Speech", "Chanson", "Opera", "Chamber Music", _
                       "Sonata", "Symphony", "Booty Bass", "Primus", "Porn Groove", "Satire", "Slow Jam", _
                       "Club", "Tango", "Samba", "Folklore", "Ballad", "Powder Ballad", "Rhythmic Soul", _
                       "Freestyle", "Duet", "Punk Rock", "Drum Solo", "A Capella", "Euro-House", "Dance Hall", _
                       "Goa", "Drum & Bass", "Club House", "Hardcore", "Terror", "Indie", "BritPop", _
                       "NegerPunk", "Polsk Punk", "Beat", "Christian Gangsta", "Heavy Metal", "Black Metal", _
                       "Crossover", "Contemporary C", "Christian Rock", "Merengue", "Salsa", "Thrash Metal", _
                       "Anime", "JPop", "SynthPop")
End Sub
