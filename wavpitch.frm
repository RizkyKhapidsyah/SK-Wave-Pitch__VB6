VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "WavPitch"
   ClientHeight    =   2940
   ClientLeft      =   2865
   ClientTop       =   1725
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2940
   ScaleWidth      =   6180
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1815
      Left            =   60
      ScaleHeight     =   1755
      ScaleWidth      =   6015
      TabIndex        =   8
      Top             =   1080
      Width           =   6075
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   3600
      TabIndex        =   6
      Text            =   ".5"
      Top             =   660
      Width           =   675
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1260
      TabIndex        =   4
      Text            =   ".5"
      Top             =   660
      Width           =   675
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Browse"
      Height          =   315
      Left            =   4920
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   900
      TabIndex        =   1
      Text            =   "c:\windows\media\Jungle Windows Start.wav"
      Top             =   60
      Width           =   3915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Play"
      Height          =   495
      Left            =   4920
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   4320
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "P&layback Multiplier:"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      Top             =   660
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "P&itch Multiplier:"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "&Filename:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const WAVE_MAPPER = -1&
Const WAVE_FORMAT_PCM = 1
Const WAVE_MAPPED = &H4
Const MMIO_READ = &H0         '  open file for reading only
Const MMIO_ALLOCBUF = &H10000     '  mmioOpen() should allocate a buffer
Const MMIO_FINDRIFF = &H20    '  mmioDescend(): find a LIST chunk
Const MMIO_FINDCHUNK = &H10    '  mmioDescend(): find a chunk by ID
Const WAVE_FORMAT_QUERY = &H1
Const MAXPNAMELEN = 32  '  max product name length (including NULL)
Const WAVECAPS_PITCH = &H1         '  supports pitch control
Const WAVECAPS_PLAYBACKRATE = &H2         '  supports playback rate control

Private Type WAVEFORMATEX
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
End Type

Private Type WAVEFORMAT
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
End Type

Private Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type

Private Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    Reserved As Long
End Type

Private Type WAVEOUTCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
    dwFormats As Long
    wChannels As Integer
    dwSupport As Long
End Type

Private Declare Function waveOutOpen Lib "winmm.dll" _
    (lphWaveOut As Long, ByVal uDeviceID As Long, _
    lpFormat As Any, ByVal dwCallback As Long, _
    ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Private Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" _
    (ByVal szFileName As String, lpmmioinfo As Any, ByVal dwOpenFlags As Long) As Long
Private Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Private Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" _
    (ByVal sz As String, ByVal uFlags As Long) As Long
Private Declare Function mmioDescend Lib "winmm.dll" _
    (ByVal hmmio As Long, lpck As Any, lpckParent As Any, ByVal uFlags As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, pch As Any, ByVal cch As Long) As Long
Private Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare Function waveOutPrepareHeader Lib "winmm.dll" _
    (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutWrite Lib "winmm.dll" _
    (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutSetPitch Lib "winmm.dll" _
    (ByVal hWaveOut As Long, ByVal dwPitch As Long) As Long
Private Declare Function waveOutSetPlaybackRate Lib "winmm.dll" _
    (ByVal hWaveOut As Long, ByVal dwRate As Long) As Long
Private Declare Function waveOutSetVolume Lib "winmm.dll" _
    (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Private Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" _
    (ByVal uDeviceID As Long, lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long

' Global Memory Flags
Const GMEM_FIXED = &H0
Const GMEM_ZEROINIT = &H40
Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Dim hmmio As Long   '// Device handle for file I/O
Dim hFormat As Long
Dim PTemp As Long
Dim WavHdle As Long
Dim bSup1 As Boolean
Dim bSup2 As Boolean

Private Sub Command1_Click()
    Call CleanUp

    Dim lRet As Long
    '// This function reads a WAVE file with the name "filename" into the memory and
    '// returns a memory handler to the data.

    Dim hSound As Long, filename As String
    filename = Text1(0).Text
    If Len(Dir$(filename)) = 0 Then MsgBox "File " & filename & " NOT found.": Exit Sub

    Dim mmckinfoParent As MMCKINFO  '// File identifier
    Dim mmckinfoSubchunk As MMCKINFO    '// Format identifier

    '// Opens the file I/O device
    hmmio = mmioOpen(filename, ByVal 0, MMIO_READ Or MMIO_ALLOCBUF)
    If hmmio = 0 Then MsgBox "1. mmioOpen failed": Exit Sub

    '// Seaks for the 'WAVE' sequency in the file
    mmckinfoParent.fccType = mmioStringToFOURCC("WAVE", 0)
    
    lRet = mmioDescend(hmmio, mmckinfoParent, ByVal 0, MMIO_FINDRIFF)
    If lRet Then MsgBox "2. mmioDescend failed": GoTo Finish

    '// Seaks for the WAVE file format
    mmckinfoSubchunk.ckid = mmioStringToFOURCC("fmt ", 0)
    lRet = mmioDescend(hmmio, mmckinfoSubchunk, mmckinfoParent, MMIO_FINDCHUNK)
    If lRet Then MsgBox "3. mmioDescend failed": GoTo Finish
    
    '// Allocates memory for format data structure
    Dim dwFmtSize As Long
    dwFmtSize = mmckinfoSubchunk.ckSize
    hFormat = GlobalAlloc(GPTR, dwFmtSize)
    
    '// Locks memory for format structure
    'WAVEFORMAT *pFormat = (WAVEFORMAT *) LocalLock(hFormat);
    Dim pFormat As WAVEFORMATEX
    lRet = mmioRead(hmmio, ByVal hFormat, dwFmtSize)
    If lRet <= 0 Then MsgBox "4. mmioRead failed - " & lRet: GoTo Finish
    Call CopyMemory(pFormat, ByVal hFormat, dwFmtSize)
    
    lRet = waveOutOpen(WavHdle, WAVE_MAPPER, ByVal hFormat, 0, 0, WAVE_FORMAT_QUERY)
    If lRet Then MsgBox "5. waveOutOpen failed": GoTo Finish
    
    lRet = mmioAscend(hmmio, mmckinfoSubchunk, 0)
    If lRet Then MsgBox "6. mmioAscend failed": GoTo Finish
    mmckinfoSubchunk.ckid = mmioStringToFOURCC("data", 0)
    
    lRet = mmioDescend(hmmio, mmckinfoSubchunk, mmckinfoParent, MMIO_FINDCHUNK)
    If lRet Then MsgBox "7. mmioDescend failed": GoTo Finish
    
    Dim dwDataSize As Long
    dwDataSize = mmckinfoSubchunk.ckSize
    
    Dim wBlockSize As Long
    wBlockSize = pFormat.nBlockAlign
    
    PTemp = GlobalAlloc(GPTR, dwDataSize)

    lRet = mmioRead(hmmio, ByVal PTemp, dwDataSize)
    If lRet <= 0 Then MsgBox "8. mmioRead failed - " & lRet: GoTo Finish

    lRet = waveOutOpen(WavHdle, WAVE_MAPPER, ByVal hFormat, 0, 0, 0)
    If lRet <> 0 Or WavHdle = 0 Then MsgBox "9. waveOutOpen failed": GoTo Finish
    
    Dim CAPS As WAVEOUTCAPS
    lRet = waveOutGetDevCaps(WavHdle, CAPS, Len(CAPS))
    If lRet Then
        MsgBox "10. waveOutGetDevCaps failed - " & lRet
    Else
        If CAPS.dwSupport And WAVECAPS_PITCH = 0 Then
            If Not bSup1 Then
                MsgBox "11. waveOutPitch is NOT supported by your hardware.  I will attempt to manually set the pitch."
                bSup1 = True
            End If
        End If
        If CAPS.dwSupport And WAVECAPS_PLAYBACKRATE = 0 Then
            If Not bSup2 Then
                MsgBox "12. waveOutSetPlaybackRate is NOT supported by your hardware.  I will attempt to manually set the playbackrate."
                bSup2 = True
            End If
        End If
    End If
    
    ReDim b(1 To dwDataSize) As Byte
    Dim i As Long, k As Long, j As Long, s As Single
    Call CopyMemory(b(1), ByVal PTemp, dwDataSize)
    
    Dim lPitch As Long, sPitch As Single
    If Not bSup1 Then
        lPitch = CLng(&H10000) * CLng(Val(Text1(1).Text))
        lRet = waveOutSetPitch(WavHdle, lPitch)
        If lRet Then
            If lRet = 8 Then
                MsgBox "11. waveOutPitch is NOT supported by your hardware."
                bSup1 = True
            Else
                MsgBox "13. waveOutSetPitch failed - " & lRet
            End If
        End If
    End If
    Picture1.Cls
    Dim x As Long, y As Long
    Dim lPlay As Long, lStep As Long, lSize As Long, sStep As Single
    lStep = dwDataSize \ Picture1.Width
    If lStep < 1 Then lStep = 1
    For i = 1 To dwDataSize Step lStep
        j = b(i)
        x = (Picture1.Width / dwDataSize) * i
        y = (Picture1.Height / 255) * j
        Picture1.PSet (x, y), 0
    Next
    
    If Not bSup2 Then
        lPlay = CLng(&H10000) * CLng(Val(Text1(2).Text))
        lRet = waveOutSetPlaybackRate(WavHdle, lPlay)
        If lRet Then
            If lRet = 8 Then
                MsgBox "12. waveOutSetPlaybackRate is NOT supported by your hardware.  I will attempt to manually set the playbackrate."
                bSup2 = True
                GoTo ManualPlayback
            Else
                MsgBox "14. waveOutSetPlaybackRate failed - " & lRet
            End If
        End If
    Else
ManualPlayback:
        If Val(Text1(2).Text) <> 1 Then
            Dim lBytesPerAllChannelSample As Long
            Dim lTotalSampleElements As Long
            Dim lNewTotalSampleElements As Long
            Dim lNewdwDataSize As Long
            
            lBytesPerAllChannelSample = (pFormat.nChannels * (pFormat.wBitsPerSample \ 8))
            lTotalSampleElements = dwDataSize \ lBytesPerAllChannelSample
            ReDim aDat(1 To lTotalSampleElements) As String
            k = 1
            For i = 1 To dwDataSize Step lBytesPerAllChannelSample
                aDat(k) = Space$(lBytesPerAllChannelSample)
                Call CopyMemory(ByVal aDat(k), b(i), lBytesPerAllChannelSample)
                k = k + 1
            Next
            lNewTotalSampleElements = lTotalSampleElements * Val(Text1(2).Text) + 32
            lNewdwDataSize = lNewTotalSampleElements * lBytesPerAllChannelSample
            ReDim b(1 To lNewdwDataSize) As Byte
            sStep = 1 / Val(Text1(2).Text)
            k = 1
            For s = 1 To lTotalSampleElements Step sStep
                If s > lTotalSampleElements Then Exit For
                For j = 1 To lBytesPerAllChannelSample
                    b(k) = Asc(Mid$(aDat(Int(s)), j, 1))
                    k = k + 1
                Next
            Next
            dwDataSize = lNewdwDataSize
            lRet = GlobalFree(PTemp)
            PTemp = GlobalAlloc(GPTR, dwDataSize)
            Call CopyMemory(ByVal PTemp, b(1), dwDataSize)
        End If
    End If
    
    '// Prepare header for sound data
    Dim WAVEHDR As WAVEHDR
    WAVEHDR.lpData = PTemp
    WAVEHDR.dwBufferLength = dwDataSize
    
    'lRet = waveOutSetVolume(WavHdle, -1)
    
    lRet = waveOutPrepareHeader(WavHdle, WAVEHDR, Len(WAVEHDR))
    If lRet Then MsgBox "15. waveOutPrepareHeader failed - " & lRet: GoTo Finish
    
    lRet = waveOutWrite(WavHdle, WAVEHDR, Len(WAVEHDR))
    If lRet Then MsgBox "16. waveOutWrite failed - " & lRet: GoTo Finish
        
    Exit Sub
Finish:
    Call CleanUp
End Sub

Private Sub Command2_Click()
    CMD.DialogTitle = "Select WAV File"
    CMD.FilterIndex = 1
    CMD.Filter = "WAV Files|*.wav"
    CMD.ShowOpen
    If Len(CMD.filename) Then
        Text1(0).Text = CMD.filename
        Command1_Click
    End If
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call CleanUp
End Sub

Sub CleanUp()
    On Local Error Resume Next
    Dim lRet As Long
    If hmmio Then lRet = mmioClose(hmmio, 0): hmmio = 0
    If WavHdle Then
        lRet = waveOutReset(WavHdle)
        lRet = waveOutClose(WavHdle)
        WavHdle = 0
    End If
    If hFormat Then lRet = GlobalFree(hFormat): hFormat = 0
    If PTemp Then lRet = GlobalFree(PTemp): PTemp = 0
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub
