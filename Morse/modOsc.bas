Attribute VB_Name = "modOsc"
Option Explicit
'--------------------------------------------------------------
' S I M P L E S T     O S C I L L A T O R   v 1.0
' Author:           Harshad Sharma
' Description:      Code to allow you to use the soundcard's output
'                   just the way you use the beep() API call to generate
'                   soound.
' NOTE:             I have used the below mentioned code and
'                   modified it to suit my needs. I do NOT
'                   claim that the Oscillator code is mine.
'                   The Link to the original code (where I got it)
'                   is given below.
'--------------------------------------------------------------
' Original Code:
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=39362&lngWId=1
'--------------------------------------------------------------

' If you have trouble check this out - you need to add the DirectX 7 to the
' References . Check if you do have dx7vb.dll in your "Windows\system32" directory
' If yes, then reference it (in the menu: Project > References) if no, then, well
' Install DX 7!!

' This API call is used to run the inbuilt timer...
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
' This TYPE is necessary if we want to get the System Time through the API
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Private MSec As Integer         ' Used for counting / storing Micro Seconds
Private MilliSec As Integer     ' Used for counting / storing Milli Seconds
Private SysTime As SYSTEMTIME   ' Used to store the System Time
Private RunTimer As Boolean     ' Used to tell the software to exit timer loop
                                ' from outside procedure. 'If RunTimer = False Then Exit Loop'

Const nSamples = 44100
Const nBasicBufferSize = 4096
Const pi = 3.14159265358979

Dim DX7 As New DirectX7, DS As DirectSound, DSB As DirectSoundBuffer
Dim PCM As WAVEFORMATEX, DSBD As DSBUFFERDESC
Dim nFreq&, nMod!, nModDir%

Public Sub InitializeBeep(hWnd As Long)
    ' Hey Guys & Gals (I hope some Gals program! ;-) please don't ask me what do
    ' the lines below mean... even I don't know, but you can contact the guy
    ' who wrote them (see top for the link)
    ' Although most is self-explanatory, please do not make any assumptions.
    
   
    nMod = 1
    Set DS = DX7.DirectSoundCreate(vbNullString)
    DS.SetCooperativeLevel hWnd, DSSCL_NORMAL
    PCM.nFormatTag = WAVE_FORMAT_PCM
    PCM.nChannels = 1
    PCM.lSamplesPerSec = nSamples
    PCM.nBitsPerSample = 8
    PCM.nBlockAlign = 1
    PCM.lAvgBytesPerSec = PCM.lSamplesPerSec * PCM.nBlockAlign
    DSBD.lFlags = DSBCAPS_STATIC
End Sub

' the name dBeep comes from: delay and beep
' (the original code does not include the delay function)
Public Sub dBeep(aFrequency As Long, aDuration As Long, Optional aVolume As Long)
    ' to make things simple, the volume parameter is kept optional,
    ' so when changing over from old code, you don't have to add the parameter
    ' unnecessarily. Only when you really need it - add it.
    
    If aVolume = 0 Then aVolume = 100   ' As the VOLUME parameter is optional,
                                        ' we have to make it 100 when it is
                                        ' not given i.e. zero
                                        
    If aVolume = -1 Then aVolume = 0    ' But we may sometimes want to keep the
                                        ' volume zero... well we have to provide
                                        ' it as -1 and it will be set to zero.
    
    ' Now adjust the frequency to the specifications required by the oscillator function
    nFreq = 1 + aFrequency * 22.049! * Log(1 + aFrequency / 1000) / Log(2)
    
    ' now call the function and play the sound
    SinBuffer nFreq / 4, 100 / 1000, 0
    
    ' let the OS perform other tasks related to our app.
    ' like telling DirectX to play the sound
    DoEvents
    
    Do While MSec < aDuration
        ' this is the timer function...
        ' all it does it wait for the system clock to show
        ' increment in the time. till that time, it gives
        ' other events more priority
        DoEvents
        Call GetSystemTime(SysTime)     ' get the system time
        If Abs(SysTime.wMilliseconds - MilliSec) >= 1 Then
            MilliSec = SysTime.wMilliseconds
            MSec = MSec + 1
        End If
        
        ' This method is NOT very accurate, nor is it very efficient
        ' It also consumes a lot of system time.
        ' But it is enough for our use.
    Loop
    MSec = 0
    
    ' set the volume to zero
    SinBuffer nFreq, 0 / 1000, 0
    
    ' I am using this stupid idea because I do not know how to stop DirectX from
    ' making any noise... if you happen to know, please contact me immediately:
    ' My mail address is given above. Your name shall be mentioned wherever necessary.
    
    ' In case of erronous loop, the DoEvents statement allows us to safely close
    ' the application rather than use the famous
    ' "Ctrl-Alt-Del" trio - (The Good, The Bad, The Ugly)!!
    ' Because it does not hang the application... with every iteration, it tells
    ' Windows to pay attention to other events related to our app.
    DoEvents
End Sub

' the code below is a stripped down version of what I had got on PSC (see comments on top for details)
' I cannot comment this code because, frankly, most of it goes way over my head
' Basically I have not tried my hand at DirectX - when I will, I'll bring out
' some nice apps on that topic.... but till that time - use this! ;-)
Private Sub SinBuffer(ByVal nFrequency&, ByVal nVolume!, Optional ByVal bSquare As Boolean)
    Dim lpBuffer() As Byte, I&, C!, nBuffer&
    C = nSamples / nFrequency
    nBuffer = (nBasicBufferSize \ C) * C
    If nBuffer = 0 Then nBuffer = C
    ReDim lpBuffer(nBuffer - 1)
    For I = 0 To nBuffer - 1
        C = Sin(I * 2 * pi / nSamples * nFrequency)
        If bSquare Then
            C = Sgn(C)
            If C = 0 Then C = 1
        End If
            lpBuffer(I) = (C * nMod * nVolume + 1) * 127.5!
    Next
    If DSBD.lBufferBytes <> nBuffer Then
        DSBD.lBufferBytes = nBuffer
        Set DSB = DS.CreateSoundBuffer(DSBD, PCM)
    End If
    
    DSB.WriteBuffer 0, 0, lpBuffer(0), DSBLOCK_ENTIREBUFFER
    DSB.Play DSBPLAY_LOOPING
End Sub

