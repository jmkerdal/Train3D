VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDirectSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' **********************************
' DirectSound Class Object
' Kerdal Jean-Michel 1999-2000
' http://members.aol.com/DemiCitron/
' Email: jeanmichel.kerdal@free.fr
' **********************************
'
Option Explicit
Private Type TypeSound
    SoundPlay As DirectSoundBuffer
    FileName As String
End Type
Private Sound() As TypeSound ' Sounds and filenames
Private SoundBuffer() As DirectSoundBuffer ' Sound buffer already playing
Private DSound As DirectSound ' DirectX sound object
Private dX7 As New DirectX7 ' DirectX7 object
Private Status As CONST_DSBSTATUSFLAGS
Private Sound3D As DirectSound3DBuffer
Public FoundCard As Boolean
'
' Virtual ear
'
Private dsPrimaryBuffer As DirectSoundBuffer
Public dsListener As DirectSound3DListener

'
' *************************
' Play a 2D sound from file
' *************************
'
Public Sub PlayFile(d$)
    If FoundCard = False Then Exit Sub
    Call LoadWave(d$, Sound(0).SoundPlay)
    Sound(0).FileName = d$
    Call Play3D(0, 0, 0, 0, DSBPLAY_DEFAULT)
End Sub

'
' *******************************
' Create a sound buffer from file
' *******************************
'
Private Sub LoadWave(SoundFile$, Play As DirectSoundBuffer)
    If FoundCard = False Then Exit Sub
    Dim BufferDesc As DSBUFFERDESC
    Dim WaveFormat As WAVEFORMATEX
    BufferDesc.lFlags = DSBCAPS_CTRL3D Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    Set Play = DSound.CreateSoundBufferFromFile(SoundFile$, BufferDesc, WaveFormat)
End Sub

'
' ***************************************************
' Add a sound in memory, return the index in the list
' ***************************************************
'
Function Add%(a$)
    If FoundCard = False Then Exit Function
    Dim i%
    If a$ = "" Then Exit Function
    For i% = 1 To UBound(Sound())
        If a$ = Sound(i%).FileName Then
            ' ***** Already done
            Add% = i%
            Exit Function
        End If
    Next i%
    Add% = UBound(Sound()) + 1
    ReDim Preserve Sound(Add%) As TypeSound
    Sound(Add%).FileName = a$
    Call LoadWave(a$, Sound(Add%).SoundPlay)
End Function

'
' ***************
' Play a 3D sound
' ***************
'
Public Function Play3D%(n%, X!, Y!, z!, Looping As CONST_DSBPLAYFLAGS, Optional Vx!, Optional Vy!, Optional Vz!)
    If FoundCard = False Then Exit Function
    Dim i%
    '
    If Sound(n%).FileName = "" Then Exit Function ' Is the sound exist?
    '
    ' ***** Copy the sound to the 2D buffer
    '
    For i% = 1 To UBound(SoundBuffer())
        Status = SoundBuffer(i%).GetStatus
        If Status = 0 Then
            Play3D% = i%
            Exit For
        End If
    Next i%
    If Play3D% = 0 Then
        Play3D% = UBound(SoundBuffer()) + 1
        ReDim Preserve SoundBuffer(Play3D%) As DirectSoundBuffer
    End If
    '
    ' ***** Get a 3D copy and play it
    '
    Set Sound3D = Sound(n%).SoundPlay.GetDirectSound3DBuffer
    Call Sound3D.SetVelocity(Vx!, Vy!, Vz!, DS3D_DEFERRED)
    Call Sound3D.SetPosition(X! / REDUCTION%, Y! / REDUCTION%, z! / REDUCTION%, DS3D_DEFERRED)
    Call dsListener.CommitDeferredSettings
    Set SoundBuffer(Play3D%) = DSound.DuplicateSoundBuffer(Sound(n%).SoundPlay)
    Call SoundBuffer(Play3D%).Play(Looping)
    Set Sound3D = Nothing
End Function

'
' ***************************
' Nouvelle position du son 3D
' ***************************
'
Public Sub SetPosition(n%, X!, Y!, z!)
    If FoundCard = False Then Exit Sub
    Set Sound3D = SoundBuffer(n%).GetDirectSound3DBuffer
    Call Sound3D.SetPosition(X!, Y!, z!, DS3D_IMMEDIATE)
    Set SoundBuffer(n%) = Sound3D.GetDirectSoundBuffer
    Set Sound3D = Nothing
End Sub

'
' *********************************
' Set the volume between
' 0 = 100%
' -10000 = 0%
' *********************************
'
Public Sub SetVolume(n%, v!)
    If FoundCard = False Then Exit Sub
    Call SoundBuffer(n%).SetVolume(v!)
End Sub

'
' *****************************
' Create the DirectSound object
' *****************************
'
Private Sub Class_Initialize()
    On Error Resume Next
    '
    Dim PrimDesc As DSBUFFERDESC
    Dim Format As WAVEFORMATEX
    Err.Clear
    Set DSound = dX7.DirectSoundCreate("")
    If Err.Number <> 0 Then
        If Err.Number <> DSERR_NODRIVER Then
            Call MsgBox("[" + Trim$(Str$(Err.Number)) + "] " + Err.Description, vbExclamation + vbOKOnly)
        End If
        Exit Sub
    End If
    '
    FoundCard = True
    PrimDesc.lBufferBytes = 0
    PrimDesc.lFlags = DSBCAPS_CTRL3D Or DSBCAPS_PRIMARYBUFFER
    Set dsPrimaryBuffer = DSound.CreateSoundBuffer(PrimDesc, Format)
    '
    Set dsListener = dsPrimaryBuffer.GetDirectSound3DListener()
    Call dsListener.SetPosition(0, 0, 0, DS3D_DEFERRED)
    Call dsListener.SetDopplerFactor(DS3D_DEFAULTDOPPLERFACTOR, DS3D_DEFERRED)
    Call dsListener.SetVelocity(1, 1, 1, DS3D_DEFERRED)
    Call dsListener.CommitDeferredSettings
    '
    ReDim SoundBuffer(0) As DirectSoundBuffer
    ReDim Sound(0) As TypeSound
End Sub

'
' *************************
' Unload DirectSound object
' *************************
'
Private Sub Class_Terminate()
    If FoundCard = False Then Exit Sub
    Dim i%
    For i% = 0 To UBound(SoundBuffer())
        Set SoundBuffer(i%) = Nothing
    Next i%
    For i% = 0 To UBound(Sound())
        Set Sound(i%).SoundPlay = Nothing
    Next i%
    Set dsListener = Nothing
    Set dsPrimaryBuffer = Nothing
    Set DSound = Nothing
    Set dX7 = Nothing
End Sub

'
' *********************************
' Attach the sound object to a form
' and set the output format
' *********************************
'
Public Sub MainForm(Form_hWnd&, Channel As Integer, Bits As Integer)
    If FoundCard = False Then Exit Sub
    Dim Format As WAVEFORMATEX
    Call DSound.SetCooperativeLevel(Form_hWnd&, DSSCL_PRIORITY)
    Format.nFormatTag = WAVE_FORMAT_PCM
    Format.nChannels = Channel%
    Format.lSamplesPerSec = 22050
    Format.nBitsPerSample = Bits%
    Format.nBlockAlign = Format.nBitsPerSample / 8 * Format.nChannels
    Format.lAvgBytesPerSec = Format.lSamplesPerSec * Format.nBlockAlign
    Call dsPrimaryBuffer.SetFormat(Format)
End Sub

'
' *********************
' Stop playing a buffer
' *********************
'
Public Sub StopPlaying(n%)
    If FoundCard = False Then Exit Sub
    Call SoundBuffer(n%).Stop
End Sub

