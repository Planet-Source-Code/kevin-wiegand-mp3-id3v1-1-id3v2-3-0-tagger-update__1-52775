VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMP3HeadReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private fstrMPEG As String
Private fbytLayer As Byte
Private fbolIsProtected As Boolean
Private fintBitrate As Integer
Private flngFrequency As Long
Private fbolIsPrivate As Boolean
Private fstrMode As String
Private fbolIsCopyright As Boolean
Private flngLength As Long

Public Property Let MPEG(ByVal strMPEG As String)
    fstrMPEG = strMPEG
End Property

Public Property Get MPEG() As String
    MPEG = fstrMPEG
End Property

Public Property Let Layer(ByVal bytLayer As Byte)
    fbytLayer = bytLayer
End Property

Public Property Get Layer() As Byte
    Layer = fbytLayer
End Property

Public Property Let IsProtected(ByVal bolIsProtected As Boolean)
    fbolIsProtected = bolIsProtected
End Property

Public Property Get IsProtected() As Boolean
    IsProtected = fbolIsProtected
End Property

Public Property Let Bitrate(ByVal intBitrate As Long)
    fintBitrate = intBitrate
End Property

Public Property Get Bitrate() As Long
    Bitrate = fintBitrate
End Property

Public Property Let Frequency(ByVal lngFrequency As Long)
    flngFrequency = lngFrequency
End Property

Public Property Get Frequency() As Long
    Frequency = flngFrequency
End Property

Public Property Let IsPrivate(ByVal bolIsPrivate As Boolean)
    fbolIsPrivate = bolIsPrivate
End Property

Public Property Get IsPrivate() As Boolean
    IsPrivate = fbolIsPrivate
End Property

Public Property Let Mode(ByVal strMode As String)
    fstrMode = strMode
End Property

Public Property Get Mode() As String
    Mode = fstrMode
End Property

Public Property Let IsCopyright(ByVal bolIsCopyright As Boolean)
    fbolIsCopyright = bolIsCopyright
End Property

Public Property Get IsCopyright() As Boolean
    IsCopyright = fbolIsCopyright
End Property

Public Property Let Length(ByVal lngLength As Long)
    flngLength = lngLength
End Property

Public Property Get Length() As Long
    Length = flngLength
End Property

Public Sub ReadMP3Header(ByVal strFileName As String)
    Dim intFile As Integer
    Dim bytByte As Byte
    Dim bytArray() As Byte
    Dim strBits As String
    Dim lngCounter As Long
    Dim lngTagSize As Long
    
    On Error GoTo ErrHan
    
    If FileExists(strFileName) = False Then Exit Sub
    
    lngTagSize = IDTagSize(strFileName)
    intFile = FreeFile
    Open strFileName For Binary Access Read Shared As intFile
        Do Until lngCounter >= lngTagSize
            Get #intFile, , bytByte
            lngCounter = lngCounter + 1
        Loop
        Do Until EOF(intFile) Or (bytByte = 255)
            Get #intFile, , bytByte
        Loop
        
        If bytByte = 255 Then
            ReDim bytArray(1 To 4)
            bytArray(1) = 255
            Get #intFile, , bytArray(2)
            Get #intFile, , bytArray(3)
            Get #intFile, , bytArray(4)
        End If
    Close
    
    strBits = NumberToBinary(bytArray(1))
    strBits = strBits & NumberToBinary(bytArray(2))
    strBits = strBits & NumberToBinary(bytArray(3))
    strBits = strBits & NumberToBinary(bytArray(4))
    
    ReadMP3Bits strBits, strFileName
    Exit Sub
    
ErrHan:
End Sub

Private Sub ReadMP3Bits(ByVal strBits As String, ByVal strFileName As String)
    Dim strSync As String
    Dim strID As String
    Dim strLayer As String
    Dim strProtected As String
    Dim strBitrate As String
    Dim strFrequency As String
    Dim strPad As String
    Dim strPrivate As String
    Dim strMode As String
    Dim strModeExtension As String
    Dim strCopyright As String
    
    If Len(strBits) <> 32 Then Exit Sub
    
    On Error GoTo ErrHan
    
    strSync = Mid(strBits, 1, 12)
    strID = Mid(strBits, 13, 1)
    strLayer = Mid(strBits, 14, 2)
    strProtected = Mid(strBits, 16, 1)
    strBitrate = Mid(strBits, 17, 4)
    strFrequency = Mid(strBits, 21, 2)
    strPad = Mid(strBits, 23, 1)
    strPrivate = Mid(strBits, 24, 1)
    strMode = Mid(strBits, 25, 2)
    strModeExtension = Mid(strBits, 27, 2)
    strCopyright = Mid(strBits, 29, 1)
    
    Select Case strLayer
        Case "01"
            Layer = 3
        Case "10"
            Layer = 2
        Case "11"
            Layer = 1
        Case Else
            Layer = 0
    End Select
    
    If strID = "1" Then
        MPEG = "MPEG-1"
        Select Case strFrequency
            Case "00"
                Frequency = 44100
            Case "01"
                Frequency = 48000
            Case "10"
                Frequency = 32000
            Case Else
                Frequency = 0
        End Select
        
        Select Case strBitrate
            Case "0001"
                Bitrate = 32
            Case "0010"
                If Layer = 1 Then
                    Bitrate = 64
                ElseIf Layer = 2 Then
                    Bitrate = 48
                ElseIf Layer = 3 Then
                    Bitrate = 40
                End If
            Case "0011"
                If Layer = 1 Then
                    Bitrate = 96
                ElseIf Layer = 2 Then
                    Bitrate = 56
                ElseIf Layer = 3 Then
                    Bitrate = 48
                End If
            Case "0100"
                If Layer = 1 Then
                    Bitrate = 128
                ElseIf Layer = 2 Then
                    Bitrate = 64
                ElseIf Layer = 3 Then
                    Bitrate = 56
                End If
            Case "0101"
                If Layer = 1 Then
                    Bitrate = 160
                ElseIf Layer = 2 Then
                    Bitrate = 80
                ElseIf Layer = 3 Then
                    Bitrate = 64
                End If
            Case "0110"
                If Layer = 1 Then
                    Bitrate = 192
                ElseIf Layer = 2 Then
                    Bitrate = 96
                ElseIf Layer = 3 Then
                    Bitrate = 80
                End If
            Case "0111"
                If Layer = 1 Then
                    Bitrate = 224
                ElseIf Layer = 2 Then
                    Bitrate = 112
                ElseIf Layer = 3 Then
                    Bitrate = 96
                End If
            Case "1000"
                If Layer = 1 Then
                    Bitrate = 256
                ElseIf Layer = 2 Then
                    Bitrate = 128
                ElseIf Layer = 3 Then
                    Bitrate = 112
                End If
            Case "1001"
                If Layer = 1 Then
                    Bitrate = 288
                ElseIf Layer = 2 Then
                    Bitrate = 160
                ElseIf Layer = 3 Then
                    Bitrate = 128
                End If
            Case "1010"
                If Layer = 1 Then
                    Bitrate = 320
                ElseIf Layer = 2 Then
                    Bitrate = 192
                ElseIf Layer = 3 Then
                    Bitrate = 160
                End If
            Case "1011"
                If Layer = 1 Then
                    Bitrate = 352
                ElseIf Layer = 2 Then
                    Bitrate = 224
                ElseIf Layer = 3 Then
                    Bitrate = 192
                End If
            Case "1100"
                If Layer = 1 Then
                    Bitrate = 384
                ElseIf Layer = 2 Then
                    Bitrate = 256
                ElseIf Layer = 3 Then
                    Bitrate = 224
                End If
            Case "1101"
                If Layer = 1 Then
                    Bitrate = 416
                ElseIf Layer = 2 Then
                    Bitrate = 320
                ElseIf Layer = 3 Then
                    Bitrate = 256
                End If
            Case "1110"
                If Layer = 1 Then
                    Bitrate = 448
                ElseIf Layer = 2 Then
                    Bitrate = 384
                ElseIf Layer = 3 Then
                    Bitrate = 320
                End If
            Case Else
                Bitrate = 0
        End Select
    Else
        MPEG = "MPEG-2"
        Select Case strFrequency
            Case "00"
                Frequency = 22050
            Case "01"
                Frequency = 24000
            Case "10"
                Frequency = 16000
            Case Else
                Frequency = 0
        End Select
        
        Select Case Bitrate
            Case "0001"
                If Layer = 1 Then
                    Bitrate = 32
                ElseIf Layer = 2 Then
                    Bitrate = 32
                ElseIf Layer = 3 Then
                    Bitrate = 8
                End If
            Case "0010"
                If Layer = 1 Then
                    Bitrate = 64
                ElseIf Layer = 2 Then
                    Bitrate = 48
                ElseIf Layer = 3 Then
                    Bitrate = 16
                End If
            Case "0011"
                If Layer = 1 Then
                    Bitrate = 96
                ElseIf Layer = 2 Then
                    Bitrate = 56
                ElseIf Layer = 3 Then
                    Bitrate = 24
                End If
            Case "0100"
                If Layer = 1 Then
                    Bitrate = 128
                ElseIf Layer = 2 Then
                    Bitrate = 64
                ElseIf Layer = 3 Then
                    Bitrate = 32
                End If
            Case "0101"
                If Layer = 1 Then
                    Bitrate = 160
                ElseIf Layer = 2 Then
                    Bitrate = 80
                ElseIf Layer = 3 Then
                    Bitrate = 64
                End If
            Case "0110"
                If Layer = 1 Then
                    Bitrate = 192
                ElseIf Layer = 2 Then
                    Bitrate = 96
                ElseIf Layer = 3 Then
                    Bitrate = 80
                End If
            Case "0111"
                If Layer = 1 Then
                    Bitrate = 224
                ElseIf Layer = 2 Then
                    Bitrate = 112
                ElseIf Layer = 3 Then
                    Bitrate = 56
                End If
            Case "1000"
                If Layer = 1 Then
                    Bitrate = 256
                ElseIf Layer = 2 Then
                    Bitrate = 128
                ElseIf Layer = 3 Then
                    Bitrate = 64
                End If
            Case "1001"
                If Layer = 1 Then
                    Bitrate = 288
                ElseIf Layer = 2 Then
                    Bitrate = 160
                ElseIf Layer = 3 Then
                    Bitrate = 128
                End If
            Case "1010"
                If Layer = 1 Then
                    Bitrate = 320
                ElseIf Layer = 2 Then
                    Bitrate = 192
                ElseIf Layer = 3 Then
                    Bitrate = 160
                End If
            Case "1011"
                If Layer = 1 Then
                    Bitrate = 352
                ElseIf Layer = 2 Then
                    Bitrate = 224
                ElseIf Layer = 3 Then
                    Bitrate = 112
                End If
            Case "1100"
                If Layer = 1 Then
                    Bitrate = 384
                ElseIf Layer = 2 Then
                    Bitrate = 256
                ElseIf Layer = 3 Then
                    Bitrate = 128
                End If
            Case "1101"
                If Layer = 1 Then
                    Bitrate = 416
                ElseIf Layer = 2 Then
                    Bitrate = 320
                ElseIf Layer = 3 Then
                    Bitrate = 256
                End If
            Case "1110"
                If Layer = 1 Then
                    Bitrate = 448
                ElseIf Layer = 2 Then
                    Bitrate = 384
                ElseIf Layer = 3 Then
                    Bitrate = 320
                End If
            Case Else
                Bitrate = 0
        End Select
    End If
    
    If strProtected = "1" Then
        IsProtected = True
    Else
        IsProtected = False
    End If
    
    If strPrivate = "1" Then
        IsPrivate = True
    Else
        IsPrivate = False
    End If
    
    Select Case strMode
        Case "00"
            Mode = "Stereo"
        Case "01"
            Mode = "Joint Stereo"
        Case "10"
            Mode = "Dual Channel"
        Case "11"
            Mode = "Mono"
        Case Else
            Mode = ""
    End Select
    
    If strCopyright = "1" Then
        IsCopyright = True
    Else
        IsCopyright = False
    End If
    
    If Bitrate <> 0 Then
        Length = ((FileLen(strFileName) - IDTagSize(strFileName)) * 8) / (Bitrate * 1000)
    Else
        Length = 0
    End If
    
    Exit Sub
    
ErrHan:
End Sub

Private Sub ClearProperties()
    MPEG = ""
    Layer = 0
    IsProtected = False
    Bitrate = 0
    Frequency = ""
    IsPrivate = False
    Mode = ""
    IsCopyright = False
End Sub

Private Function NumberToBinary(ByVal intNumber As Integer) As String
    Dim lngNumber As Long
    Dim strTemp As String
    Dim strPad As String
    
    On Error GoTo ErrHan
    
    lngNumber = intNumber
    
    Do
        If (lngNumber Mod 2) = 0 Then
            strTemp = "0" & strTemp
        Else
            strTemp = "1" & strTemp
        End If
        
        lngNumber = lngNumber \ 2
        DoEvents
    Loop Until lngNumber = 0
    
    If Len(strTemp) < 8 Then strPad = String(8 - Len(strTemp), "0")
    NumberToBinary = strPad & strTemp
    
    Exit Function
    
ErrHan:
    NumberToBinary = ""
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

Private Function IDTagSize(ByVal strFileName As String) As Long
    Dim intFile As Integer
    Dim strID3 As String * 3
    Dim bytByte As Byte
    Dim bytSize() As Byte
    Dim lngSize As Long
    
    On Error GoTo ErrHan
    
    intFile = FreeFile
    Open strFileName For Binary Access Read Shared As intFile
        Get #intFile, , strID3
        If strID3 = "ID3" Then
            ReDim bytSize(1 To 4)
            Get #intFile, , bytByte
            Get #intFile, , bytByte
            Get #intFile, , bytByte
            Get #intFile, , bytSize
            lngSize = LeftShift(bytSize(1), 21) + LeftShift(bytSize(2), 14) + LeftShift(bytSize(3), 7) + bytSize(4) + 10
        End If
    Close
    
    IDTagSize = lngSize
    
    Exit Function
    
ErrHan:
End Function

Private Function LeftShift(ByVal intNumber As Integer, ByVal bytAmount As Byte) As Long
    Dim strTemp As String
   
    If bytAmount = 0 Then
        LeftShift = intNumber
    Else
        strTemp = NumberToBinary(intNumber)
        strTemp = strTemp & String(bytAmount, "0")
        LeftShift = BinaryToLong(strTemp)
    End If
End Function

Private Function BinaryToLong(ByVal strBinary As String) As Long
    Dim lngTotal As Long
    Dim lngFactor As Long
    Dim lngCounter As Long
    Dim bytBit As Byte
    
    lngFactor = 0
    For lngCounter = Len(strBinary) To 1 Step -1
        bytBit = CByte(Mid(strBinary, lngCounter, 1))
        lngTotal = lngTotal + (bytBit * (2 ^ lngFactor))
        lngFactor = lngFactor + 1
    Next
    
    BinaryToLong = lngTotal
End Function
