Attribute VB_Name = "i_sound"
' i_sound.c - system interface for sound, rewritten for Win32
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
'
' linuxdoom mixed eight channels into a buffer for /dev/dsp and had no music.
' Here:
'  * Sound effects: the same mixer (8 channels, volume lookup tables, pitch
'    stepping) at 44100 Hz (most Freedoom effects are 22050 or 44100 Hz, the
'    DOOM ones 11025; each sound is stepped at its own rate).  It feeds a ring
'    of waveOut buffers.  VB6 has no threads, so the game loop calls
'    I_UpdateSound every frame and while it waits for the next tic; SNDLEAD
'    buffers (60 ms) are kept queued.
'  * I_StopSound, I_SoundIsPlaying and I_UpdateSoundParams work as in the DOS
'    version (linuxdoom left them as dummies, and its I_SoundIsPlaying
'    compared a handle with gametic).
'  * Music: MUS (DOOM IWADs) and standard MIDI files (Freedoom) are converted
'    to one event list with millisecond times and played with the midiStream
'    API, which times the events itself; I_UpdateMusic only refills buffers.
'    The music volume scales the channel volume controllers (midiOutSetVolume
'    would change the volume of the whole program).
' -nosound, -nosfx and -nomusic leave the devices closed.
' Test aids (no device is opened): -sndcapture <file.wav> mixes 1260 samples
' per gametic into a WAV file; -muslog <file> lists the music events.
Option Explicit

Private Type WAVEFORMATEX
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
End Type

Private Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    reserved As Long
End Type

Private Type MIDIHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    lpNext As Long
    reserved As Long
    dwOffset As Long
    dwReserved(0 To 7) As Long
End Type

Private Declare Function waveOutOpen Lib "winmm.dll" (lphWaveOut As Long, ByVal uDeviceID As Long, _
    lpFormat As WAVEFORMATEX, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hwo As Long, ByVal pwh As Long, ByVal cbwh As Long) As Long
Private Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hwo As Long, ByVal pwh As Long, ByVal cbwh As Long) As Long
Private Declare Function waveOutWrite Lib "winmm.dll" (ByVal hwo As Long, ByVal pwh As Long, ByVal cbwh As Long) As Long
Private Declare Function waveOutReset Lib "winmm.dll" (ByVal hwo As Long) As Long
Private Declare Function waveOutClose Lib "winmm.dll" (ByVal hwo As Long) As Long

Private Declare Function midiStreamOpen Lib "winmm.dll" (phms As Long, puDeviceID As Long, ByVal cMidi As Long, _
    ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function midiStreamClose Lib "winmm.dll" (ByVal hms As Long) As Long
Private Declare Function midiStreamProperty Lib "winmm.dll" (ByVal hms As Long, ByVal lppropdata As Long, ByVal dwProperty As Long) As Long
Private Declare Function midiStreamOut Lib "winmm.dll" (ByVal hms As Long, ByVal pmh As Long, ByVal cbmh As Long) As Long
Private Declare Function midiStreamRestart Lib "winmm.dll" (ByVal hms As Long) As Long
Private Declare Function midiStreamPause Lib "winmm.dll" (ByVal hms As Long) As Long
Private Declare Function midiStreamStop Lib "winmm.dll" (ByVal hms As Long) As Long
Private Declare Function midiOutPrepareHeader Lib "winmm.dll" (ByVal hmo As Long, ByVal pmh As Long, ByVal cbmh As Long) As Long
Private Declare Function midiOutUnprepareHeader Lib "winmm.dll" (ByVal hmo As Long, ByVal pmh As Long, ByVal cbmh As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hmo As Long, ByVal dwMsg As Long) As Long
Private Declare Function midiStreamPosition Lib "winmm.dll" (ByVal hms As Long, ByVal pmmt As Long, ByVal cbmmt As Long) As Long

Private Const WAVE_MAPPER As Long = -1
Private Const MIDI_MAPPER As Long = -1
Private Const WHDR_DONE As Long = 1
Private Const MHDR_DONE As Long = 1
Private Const MIDIERR_STILLPLAYING As Long = 65
Private Const MIDIPROP_SET As Long = &H80000000
Private Const MIDIPROP_TIMEDIV As Long = 1
Private Const MIDIPROP_TEMPO As Long = 2
Private Const MEVT_TEMPO As Long = &H1000000     ' dwEvent high byte 1
Private Const MEVT_NOP As Long = &H2000000       ' dwEvent high byte 2
Private Const TIME_TICKS As Long = &H20

' ---------------------------------------------------------------- sound effects

Private Const OUTRATE As Long = 44100            ' Hz (linuxdoom: 11025)
Private Const NUM_CHANNELS As Long = 8
Private Const BUFSAMPLES As Long = 441           ' stereo frames per buffer (10 ms)
Private Const NUMBUFS As Long = 16
Private Const SNDLEAD As Long = 6                ' buffers kept queued
Private Const CAPFRAMES As Long = OUTRATE \ TICRATE

Private hwo As Long                              ' waveOut handle, 0 = none
Private sfxon As Boolean                         ' mixer running (device or capture)
Private whdr(0 To NUMBUFS - 1) As WAVEHDR
Private wqueued(0 To NUMBUFS - 1) As Boolean
Private mixbuf() As Integer                      ' NUMBUFS buffers of L,R pairs
Private wnext As Long                            ' next buffer of the ring to fill
Private nqueued As Long
Private wplayed As Long                          ' buffers written (logged at shutdown)
Private underruns As Long                        ' times the device ran out of audio

Private capfile As Integer                       ' -sndcapture file, 0 = none
Private capbuf() As Integer
Private captic As Long
Private capbytes As Long

' The channel data positions (C: pointers channels / channelsend), here
' wadbuf offsets; chPos = 0 means the channel is free.
Private chPos(0 To NUM_CHANNELS - 1) As Long
Private chEnd(0 To NUM_CHANNELS - 1) As Long
' The channel step amount and a 0.16 bit remainder of last step.
Private chStep(0 To NUM_CHANNELS - 1) As Long
Private chFrac(0 To NUM_CHANNELS - 1) As Long
' Time/gametic that the channel started playing, used to determine oldest.
Private chStart(0 To NUM_CHANNELS - 1) As Long
Private chHandle(0 To NUM_CHANNELS - 1) As Long
' SFX id of the playing sound effect, used to catch duplicates (chainsaw).
Private chId(0 To NUM_CHANNELS - 1) As Long
' Left and right volume lookups (offsets into vol_lookup).
Private chLeft(0 To NUM_CHANNELS - 1) As Long
Private chRight(0 To NUM_CHANNELS - 1) As Long

' Pitch to stepping lookup.
Private steptable(0 To 255) As Long
' Volume lookups.
Private vol_lookup(0 To 128& * 256 - 1) As Long

Private handlenums As Long

' ---------------------------------------------------------------- music

Private Const NUMMIDIBUFS As Long = 6
Private Const CHUNKEVENTS As Long = 1024         ' MIDIEVENTs (12 bytes) per buffer
Private Const CHUNKMS As Long = 500              ' at most this much music per buffer
Private Const MUSLEADMS As Long = 3000           ' -muslog: list this far ahead

Private hms As Long                              ' midiStream handle, 0 = none
Private musicon As Boolean                       ' device or -muslog open
Private muslog As Integer                        ' -muslog file, 0 = none
Private musvol As Long                           ' 0..15
Private chanvol(0 To 15) As Long                 ' the song's channel volumes (CC 7)

' the registered song: event times in milliseconds and the MIDI short
' messages (or MEVT_NOP / MEVT_TEMPO events)
Private evTime() As Long
Private evMsg() As Long
Private numev As Long
Private loopstart As Long                        ' first event after the reset prelude
Private songok As Boolean

' MIDI file parsing: events of all tracks in file ticks
Private rawTick() As Long
Private rawMsg() As Long
Private nraw As Long

Private musplaying As Boolean
Private musloop As Boolean
Private evpos As Long                            ' next event to stream
Private lasttime As Long                         ' time of the last streamed event
Private streamtime As Long                       ' ms of music streamed so far
Private musstart As Long                         ' -muslog: I_GetTimeMS at the start
Private mhdr(0 To NUMMIDIBUFS - 1) As MIDIHDR
Private mqueued(0 To NUMMIDIBUFS - 1) As Boolean
Private midibuf() As Long                        ' NUMMIDIBUFS * CHUNKEVENTS * 3

'
' SFX API
' Note: this was called by S_Init.
' Init internal lookups (raw data, mixing buffer, channels).
' This function sets up internal lookups used during
'  the mixing process.
'
Public Sub I_SetChannels()
    Dim i As Long, j As Long

    ' This table provides step widths for pitch parameters.
    For i = -128 To 127
        steptable(i + 128) = Int(2 ^ (i / 64#) * 65536#)
    Next

    ' Generates volume lookup tables
    '  which also turn the unsigned samples
    '  into signed samples.
    For i = 0 To 127
        For j = 0 To 255
            vol_lookup(i * 256 + j) = (i * (j - 128) * 256) \ 127
        Next
    Next
End Sub

'
' Retrieve the raw data lump index
'  for a given SFX name.
'
Public Function I_GetSfxLumpNum(ByVal sfx As Long) As Long
    Dim lump As Long

    ' a linked sound (chgun -> pistol) plays the data of its link
    If S_sfx(sfx).link >= 0 Then sfx = S_sfx(sfx).link

    ' Sounds from DOOM II can be requested by the shareware game: use a
    ' default sound for replacement (linuxdoom's getsfx).
    lump = W_CheckNumForName("ds" & S_sfx(sfx).name)
    If lump = -1 Then lump = W_GetNumForName("dspistol")
    I_GetSfxLumpNum = lump
End Function

' Sets the left and right volume lookups of a channel.
'  range of seperation is: 1 - 256
Private Sub SetChannelVolume(ByVal slot As Long, ByVal volume As Long, ByVal seperation As Long)
    Dim leftvol As Long, rightvol As Long

    seperation = seperation + 1

    ' Per left/right channel.
    '  x^2 seperation,
    '  adjust volume properly.
    leftvol = volume - (volume * seperation * seperation) \ 65536
    seperation = seperation - 257
    rightvol = volume - (volume * seperation * seperation) \ 65536

    ' Sanity check, clamp volume (the C bombed out with I_Error).
    If rightvol < 0 Then rightvol = 0
    If rightvol > 127 Then rightvol = 127
    If leftvol < 0 Then leftvol = 0
    If leftvol > 127 Then leftvol = 127

    ' Get the proper lookup table piece
    '  for this volume level
    chLeft(slot) = leftvol * 256
    chRight(slot) = rightvol * 256
End Sub

'
' This function adds a sound to the
'  list of currently active sounds,
'  which is maintained as a given number
'  (eight, usually) of internal channels.
' Returns a handle.
'
Private Function addsfx(ByVal sfxid As Long, ByVal volume As Long, ByVal stp As Long, ByVal seperation As Long) As Long
    Dim i As Long, rc As Long, oldest As Long, oldestnum As Long, slot As Long
    Dim lump As Long, p As Long, size As Long, rate As Long, cnt As Long

    ' Chainsaw troubles.
    ' Play these sound effects only one at a time.
    If sfxid = sfx_sawup Or sfxid = sfx_sawidl Or sfxid = sfx_sawful Or _
       sfxid = sfx_sawhit Or sfxid = sfx_stnmov Or sfxid = sfx_pistol Then
        ' Loop all channels, check.
        For i = 0 To NUM_CHANNELS - 1
            ' Active, and using the same SFX?
            If chPos(i) <> 0 And chId(i) = sfxid Then
                ' Reset.
                chPos(i) = 0
                ' We are sure that iff,
                '  there will only be one.
                Exit For
            End If
        Next
    End If

    ' Loop all channels to find oldest SFX.
    oldest = gametic
    oldestnum = 0
    For i = 0 To NUM_CHANNELS - 1
        If chPos(i) = 0 Then Exit For
        If chStart(i) < oldest Then
            oldestnum = i
            oldest = chStart(i)
        End If
    Next

    ' If we found a channel, fine.
    ' If not, we overwrite the oldest one.
    If i = NUM_CHANNELS Then
        slot = oldestnum
    Else
        slot = i
    End If

    ' Reset current handle number, limited to 100..65535.
    If handlenums < 100 Or handlenums > 65535 Then handlenums = 100

    ' Assign current handle number.
    rc = handlenums
    handlenums = handlenums + 1
    addsfx = rc

    If Not sfxon Then Exit Function

    ' The raw data: a DMX sound lump (format 3, sample rate, sample count,
    ' then unsigned 8 bit samples).
    lump = S_sfx(sfxid).lumpnum
    If lump < 0 Then Exit Function
    p = W_CacheLumpNum(lump)
    size = W_LumpLength(lump)
    If size <= 8 Then Exit Function
    If wadbuf(p) <> 3 Or wadbuf(p + 1) <> 0 Then Exit Function
    rate = wadbuf(p + 2) + wadbuf(p + 3) * 256&
    CopyMemory cnt, wadbuf(p + 4), 4
    If cnt > size - 8 Or cnt < 0 Then cnt = size - 8
    If rate = 0 Or cnt = 0 Then Exit Function

    ' Okay, in the less recent channel,
    '  we will handle the new SFX.
    ' Set pointer to raw data and to the end of it.
    chPos(slot) = p + 8
    chEnd(slot) = p + 8 + cnt

    chHandle(slot) = rc

    ' Set stepping (scaled from the sound's own rate to OUTRATE).
    chStep(slot) = CLng(CDbl(stp) * rate / OUTRATE)
    chFrac(slot) = 0
    ' Should be gametic, I presume.
    chStart(slot) = gametic

    SetChannelVolume slot, volume, seperation

    ' Preserve sound SFX id,
    '  e.g. for avoiding duplicates of chainsaw.
    chId(slot) = sfxid
End Function

'
' Starting a sound means adding it
'  to the current list of active sounds
'  in the internal channels.
' As our sound handling does not handle
'  priority, it is ignored.
'
Public Function I_StartSound(ByVal id As Long, ByVal vol As Long, ByVal sep As Long, _
        ByVal pitch As Long, ByVal priority As Long) As Long
    I_StartSound = addsfx(id, vol, steptable(pitch), sep)
End Function

Public Sub I_StopSound(ByVal handle As Long)
    Dim i As Long
    For i = 0 To NUM_CHANNELS - 1
        If chPos(i) <> 0 And chHandle(i) = handle Then chPos(i) = 0
    Next
End Sub

Public Function I_SoundIsPlaying(ByVal handle As Long) As Boolean
    Dim i As Long
    For i = 0 To NUM_CHANNELS - 1
        If chPos(i) <> 0 And chHandle(i) = handle Then
            I_SoundIsPlaying = True
            Exit Function
        End If
    Next
End Function

' Volume and separation follow a moving source.  The pitch stays as it was
' started (S_UpdateSounds passes the unperturbed pitch).
Public Sub I_UpdateSoundParams(ByVal handle As Long, ByVal vol As Long, ByVal sep As Long, ByVal pitch As Long)
    Dim i As Long
    For i = 0 To NUM_CHANNELS - 1
        If chPos(i) <> 0 And chHandle(i) = handle Then SetChannelVolume i, vol, sep
    Next
End Sub

'
' This function loops all active (internal) sound
'  channels, retrieves a given number of samples
'  from the raw sound data, modifies it according
'  to the current (internal) channel parameters,
'  mixes the per channel samples into dst(o...),
'  clamping it to the allowed range.
'
Private Sub MixSamples(dst() As Integer, ByVal o As Long, ByVal frames As Long)
    Dim n As Long, chan As Long, sample As Long, dl As Long, dr As Long

    For n = 1 To frames
        ' Reset left/right value.
        dl = 0
        dr = 0

        For chan = 0 To NUM_CHANNELS - 1
            ' Check channel, if active.
            If chPos(chan) <> 0 Then
                ' Get the raw data from the channel.
                sample = wadbuf(chPos(chan))
                ' Add left and right part
                '  for this channel (sound)
                '  to the current data.
                ' Adjust volume accordingly.
                dl = dl + vol_lookup(chLeft(chan) + sample)
                dr = dr + vol_lookup(chRight(chan) + sample)
                ' Increment index
                chFrac(chan) = chFrac(chan) + chStep(chan)
                ' MSB is next sample
                chPos(chan) = chPos(chan) + (chFrac(chan) \ 65536)
                ' Limit to LSB
                chFrac(chan) = chFrac(chan) And &HFFFF&
                ' Check whether we are done.
                If chPos(chan) >= chEnd(chan) Then chPos(chan) = 0
            End If
        Next

        ' Clamp to range.
        If dl > 32767 Then
            dl = 32767
        ElseIf dl < -32768 Then
            dl = -32768
        End If
        If dr > 32767 Then
            dr = 32767
        ElseIf dr < -32768 Then
            dr = -32768
        End If

        dst(o) = dl
        dst(o + 1) = dr
        o = o + 2
    Next
End Sub

'
' Mixes new audio and hands it to the sound device (C: I_UpdateSound mixed,
' I_SubmitSound wrote the buffer).  Called every frame and while the game
' loop waits for a tic.
'
Public Sub I_UpdateSound()
    Dim b As Long

    If capfile <> 0 Then
        ' -sndcapture: one tic of audio per gametic
        Do While captic < gametic
            MixSamples capbuf, 0, CAPFRAMES
            Put #capfile, , capbuf
            capbytes = capbytes + CAPFRAMES * 4
            captic = captic + 1
        Loop
        Exit Sub
    End If

    If hwo = 0 Then Exit Sub

    ' take back the buffers the device has played
    For b = 0 To NUMBUFS - 1
        If wqueued(b) Then
            If (whdr(b).dwFlags And WHDR_DONE) <> 0 Then
                wqueued(b) = False
                nqueued = nqueued - 1
            End If
        End If
    Next
    If nqueued = 0 And wplayed > 0 Then underruns = underruns + 1

    ' keep SNDLEAD buffers queued
    Do While nqueued < SNDLEAD
        b = wnext
        If wqueued(b) Then Exit Do
        MixSamples mixbuf, b * BUFSAMPLES * 2, BUFSAMPLES
        whdr(b).dwFlags = whdr(b).dwFlags And Not WHDR_DONE
        If waveOutWrite(hwo, VarPtr(whdr(b)), LenB(whdr(b))) <> 0 Then Exit Do
        wqueued(b) = True
        nqueued = nqueued + 1
        wplayed = wplayed + 1
        wnext = (wnext + 1) Mod NUMBUFS
    Loop
End Sub

Private Sub PutLong(ByVal v As Long)
    Put #capfile, , v
End Sub

Private Sub PutShort(ByVal v As Integer)
    Put #capfile, , v
End Sub

Private Sub PutTag(ByVal s As String)
    Put #capfile, , s
End Sub

Private Sub WriteWavHeader()
    Seek #capfile, 1
    PutTag "RIFF"
    PutLong 36 + capbytes
    PutTag "WAVEfmt "
    PutLong 16
    PutShort 1                                   ' PCM
    PutShort 2                                   ' stereo
    PutLong OUTRATE
    PutLong OUTRATE * 4
    PutShort 4
    PutShort 16
    PutTag "data"
    PutLong capbytes
End Sub

Public Sub I_InitSound()
    Dim wfx As WAVEFORMATEX, b As Long, rc As Long, p As Long

    p = M_CheckParm("-sndcapture")
    If p > 0 And p < myargc - 1 Then
        If Len(Dir$(myargv(p + 1))) > 0 Then Kill myargv(p + 1)
        capfile = FreeFile
        Open myargv(p + 1) For Binary Access Write As #capfile
        ReDim capbuf(0 To CAPFRAMES * 2 - 1)
        capbytes = 0
        captic = gametic
        WriteWavHeader
        sfxon = True
        I_Printf "I_InitSound: mixing into " & myargv(p + 1)
        Exit Sub
    End If

    If M_CheckParm("-nosound") <> 0 Or M_CheckParm("-nosfx") <> 0 Then
        I_Printf "I_InitSound: sound effects off"
        Exit Sub
    End If

    ' Secure and configure sound device first.
    wfx.wFormatTag = 1                           ' WAVE_FORMAT_PCM
    wfx.nChannels = 2
    wfx.nSamplesPerSec = OUTRATE
    wfx.wBitsPerSample = 16
    wfx.nBlockAlign = 4
    wfx.nAvgBytesPerSec = OUTRATE * 4
    wfx.cbSize = 0
    rc = waveOutOpen(hwo, WAVE_MAPPER, wfx, 0, 0, 0)
    If rc <> 0 Then
        hwo = 0
        I_Printf "I_InitSound: could not open the sound device (error " & rc & ")"
        Exit Sub
    End If

    ' the mixing buffers stay at one address while the device owns them
    ReDim mixbuf(0 To NUMBUFS * BUFSAMPLES * 2 - 1)
    For b = 0 To NUMBUFS - 1
        whdr(b).lpData = VarPtr(mixbuf(b * BUFSAMPLES * 2))
        whdr(b).dwBufferLength = BUFSAMPLES * 4
        whdr(b).dwFlags = 0
        waveOutPrepareHeader hwo, VarPtr(whdr(b)), LenB(whdr(b))
        wqueued(b) = False
    Next
    wnext = 0
    nqueued = 0
    sfxon = True

    I_Printf "I_InitSound: sound module ready (" & OUTRATE & " Hz)"
End Sub

Public Sub I_ShutdownSound()
    Dim b As Long

    If hwo <> 0 Then
        I_Printf "I_ShutdownSound: " & wplayed & " buffers played, " & underruns & " underruns"
        waveOutReset hwo
        For b = 0 To NUMBUFS - 1
            waveOutUnprepareHeader hwo, VarPtr(whdr(b)), LenB(whdr(b))
            wqueued(b) = False
        Next
        waveOutClose hwo
        hwo = 0
    End If

    If capfile <> 0 Then
        WriteWavHeader
        Close #capfile
        capfile = 0
    End If

    sfxon = False
End Sub

'
' MUSIC API.
'

' A channel volume (0..127) scaled by the music volume.
Private Function ScaleVol(ByVal v As Long) As Long
    ScaleVol = (v * musvol) \ 15
End Function

Private Sub AddEvent(ByVal t As Long, ByVal msg As Long)
    If numev > UBound(evTime) Then
        ReDim Preserve evTime(0 To 2 * numev - 1)
        ReDim Preserve evMsg(0 To 2 * numev - 1)
    End If
    evTime(numev) = t
    evMsg(numev) = msg
    numev = numev + 1
End Sub

Private Sub AddRaw(ByVal tick As Long, ByVal msg As Long)
    If nraw > UBound(rawTick) Then
        ReDim Preserve rawTick(0 To 2 * nraw - 1)
        ReDim Preserve rawMsg(0 To 2 * nraw - 1)
    End If
    rawTick(nraw) = tick
    rawMsg(nraw) = msg
    nraw = nraw + 1
End Sub

' Every song starts from the same channel state; the loop starts after this.
Private Sub AddPrelude()
    Dim ch As Long

    numev = 0
    AddEvent 0, MEVT_TEMPO Or 1000000            ' 1000 ticks per second
    For ch = 0 To 15
        AddEvent 0, &HB0 Or ch Or 121 * 256&     ' reset all controllers
        AddEvent 0, &HC0 Or ch                   ' program 0
        AddEvent 0, &HB0 Or ch Or 7 * 256& Or 100 * 65536   ' volume 100
        AddEvent 0, &HB0 Or ch Or 10 * 256& Or 64 * 65536   ' pan centre
    Next
    loopstart = numev
End Sub

Private Function IsTag(ByVal p As Long, ByVal tag As String) As Boolean
    Dim i As Long
    For i = 0 To Len(tag) - 1
        If wadbuf(p + i) <> Asc(Mid$(tag, i + 1, 1)) Then Exit Function
    Next
    IsTag = True
End Function

Private Function RdU16BE(ByVal p As Long) As Long
    RdU16BE = wadbuf(p) * 256& + wadbuf(p + 1)
End Function

Private Function RdU32BE(ByVal p As Long) As Long
    RdU32BE = ((wadbuf(p) * 256& + wadbuf(p + 1)) * 256& + wadbuf(p + 2)) * 256& + wadbuf(p + 3)
End Function

' Variable length quantity (MIDI delta times, MUS delays).
Private Function ReadVLQ(q As Long, ByVal qend As Long) As Long
    Dim v As Long, b As Long, n As Long
    Do
        If q >= qend Then Exit Do
        b = wadbuf(q)
        q = q + 1
        v = v * 128 + (b And 127)
        n = n + 1
    Loop While (b And 128) <> 0 And n < 4
    ReadVLQ = v
End Function

' MUS: the DMX score format of the DOOM IWADs, 140 ticks per second.
Private Function ConvertMUS(ByVal p As Long, ByVal size As Long) As Boolean
    Dim q As Long, qend As Long, b As Long, typ As Long, ch As Long, mch As Long
    Dim note As Long, v As Long, ctl As Long, t As Long, ms As Long
    Dim vel(0 To 15) As Long

    q = p + wadbuf(p + 6) + wadbuf(p + 7) * 256&            ' score start
    qend = q + wadbuf(p + 4) + wadbuf(p + 5) * 256&         ' + score length
    If qend > p + size Then qend = p + size

    For ch = 0 To 15
        vel(ch) = 127
    Next
    t = 0

    Do While q < qend
        b = wadbuf(q)
        q = q + 1
        typ = (b \ 16) And 7
        ch = b And 15
        ' MUS channel 15 is percussion (MIDI 9); 9..14 move up one
        If ch = 15 Then
            mch = 9
        ElseIf ch >= 9 Then
            mch = ch + 1
        Else
            mch = ch
        End If
        ms = Int(t * 1000# / 140# + 0.5)

        Select Case typ
        Case 0                                   ' release note
            note = wadbuf(q) And 127
            q = q + 1
            AddEvent ms, &H80 Or mch Or note * 256
        Case 1                                   ' play note
            note = wadbuf(q)
            q = q + 1
            If (note And 128) <> 0 Then
                vel(ch) = wadbuf(q) And 127
                q = q + 1
            End If
            AddEvent ms, &H90 Or mch Or (note And 127) * 256 Or vel(ch) * 65536
        Case 2                                   ' pitch wheel (8 bits, 128 = centre)
            v = wadbuf(q) * 64&
            q = q + 1
            AddEvent ms, &HE0 Or mch Or (v And 127) * 256 Or (v \ 128) * 65536
        Case 3                                   ' system event
            ctl = wadbuf(q) And 127
            q = q + 1
            Select Case ctl
            Case 10: AddEvent ms, &HB0 Or mch Or 120 * 256&     ' all sounds off
            Case 11: AddEvent ms, &HB0 Or mch Or 123 * 256&     ' all notes off
            Case 12: AddEvent ms, &HB0 Or mch Or 126 * 256&     ' mono
            Case 13: AddEvent ms, &HB0 Or mch Or 127 * 256&     ' poly
            Case 14: AddEvent ms, &HB0 Or mch Or 121 * 256&     ' reset all controllers
            End Select
        Case 4                                   ' change controller
            ctl = wadbuf(q)
            v = wadbuf(q + 1)
            q = q + 2
            If v > 127 Then v = 127
            Select Case ctl
            Case 0: AddEvent ms, &HC0 Or mch Or v * 256           ' instrument
            Case 1: AddEvent ms, &HB0 Or mch Or 0 Or v * 65536    ' bank select
            Case 2: AddEvent ms, &HB0 Or mch Or 1 * 256& Or v * 65536     ' modulation
            Case 3: AddEvent ms, &HB0 Or mch Or 7 * 256& Or v * 65536     ' volume
            Case 4: AddEvent ms, &HB0 Or mch Or 10 * 256& Or v * 65536    ' pan
            Case 5: AddEvent ms, &HB0 Or mch Or 11 * 256& Or v * 65536    ' expression
            Case 6: AddEvent ms, &HB0 Or mch Or 91 * 256& Or v * 65536    ' reverb
            Case 7: AddEvent ms, &HB0 Or mch Or 93 * 256& Or v * 65536    ' chorus
            Case 8: AddEvent ms, &HB0 Or mch Or 64 * 256& Or v * 65536    ' sustain
            Case 9: AddEvent ms, &HB0 Or mch Or 67 * 256& Or v * 65536    ' soft pedal
            End Select
        Case 6                                   ' score end
            Exit Do
        End Select                               ' 5 (end of measure), 7: no data

        ' a delay follows the last event of a group
        If (b And 128) <> 0 Then t = t + ReadVLQ(q, qend)
    Loop

    AddEvent Int(t * 1000# / 140# + 0.5), MEVT_NOP   ' the song length, for looping
    ConvertMUS = True
End Function

' One track of a MIDI file into rawTick / rawMsg.
Private Sub ParseTrack(ByVal q As Long, ByVal qend As Long)
    Dim tick As Long, st As Long, running As Long, d1 As Long, d2 As Long
    Dim mtype As Long, mlen As Long

    Do While q < qend
        tick = tick + ReadVLQ(q, qend)
        If q >= qend Then Exit Do
        st = wadbuf(q)
        If st >= &H80 Then
            q = q + 1
        Else
            ' running status: this byte is the first data byte
            st = running
            If st = 0 Then Exit Do
        End If

        Select Case st And &HF0
        Case &H80, &H90, &HA0, &HB0, &HE0
            If q + 2 > qend Then Exit Do
            d1 = wadbuf(q) And 127
            d2 = wadbuf(q + 1) And 127
            q = q + 2
            running = st
            AddRaw tick, st Or d1 * 256 Or d2 * 65536
        Case &HC0, &HD0
            If q + 1 > qend Then Exit Do
            d1 = wadbuf(q) And 127
            q = q + 1
            running = st
            AddRaw tick, st Or d1 * 256
        Case Else
            If st = &HFF Then                    ' meta event
                If q >= qend Then Exit Do
                mtype = wadbuf(q)
                q = q + 1
                mlen = ReadVLQ(q, qend)
                If mtype = &H51 And mlen >= 3 And q + 3 <= qend Then
                    AddRaw tick, MEVT_TEMPO Or (wadbuf(q) * 65536 + wadbuf(q + 1) * 256& + wadbuf(q + 2))
                ElseIf mtype = &H2F Then         ' end of track
                    AddRaw tick, MEVT_NOP
                    Exit Do
                End If
                q = q + mlen
            ElseIf st = &HF0 Or st = &HF7 Then   ' system exclusive: skipped
                mlen = ReadVLQ(q, qend)
                q = q + mlen
            Else
                Exit Do                          ' not allowed in a file
            End If
        End Select
    Loop
End Sub

' Standard MIDI file (format 0 or 1): the tracks are merged in time order
' and the tempo map turns file ticks into milliseconds.
Private Function ConvertMIDI(ByVal p As Long, ByVal size As Long) As Boolean
    Dim lend As Long, q As Long, ntrks As Long, division As Long, tr As Long
    Dim clen As Long, best As Long, i As Long, tick As Long, lasttick As Long
    Dim msg As Long, tempo As Long
    Dim tstart() As Long, tend() As Long, ptr() As Long
    Dim mspertick As Double, t As Double, endms As Double

    lend = p + size
    ntrks = RdU16BE(p + 10)
    division = RdU16BE(p + 12)
    If ntrks < 1 Or division = 0 Then Exit Function
    q = p + 8 + RdU32BE(p + 4)

    ReDim tstart(0 To ntrks - 1)
    ReDim tend(0 To ntrks - 1)
    ReDim ptr(0 To ntrks - 1)
    ReDim rawTick(0 To 1023)
    ReDim rawMsg(0 To 1023)
    nraw = 0

    For tr = 0 To ntrks - 1
        tstart(tr) = nraw
        ' find the next track chunk
        Do While q + 8 <= lend
            clen = RdU32BE(q + 4)
            If clen < 0 Then clen = lend
            If IsTag(q, "MTrk") Then Exit Do
            q = q + 8 + clen
        Loop
        If q + 8 <= lend Then
            If q + 8 + clen > lend Then clen = lend - q - 8
            ParseTrack q + 8, q + 8 + clen
            q = q + 8 + clen
        End If
        tend(tr) = nraw
        ptr(tr) = tstart(tr)
    Next

    tempo = 500000                               ' 120 beats per minute
    If (division And &H8000&) <> 0 Then
        ' SMPTE: frames per second * ticks per frame
        mspertick = 1000# / ((256 - division \ 256) * (division And 255))
    End If

    Do
        ' the track with the earliest next event (the first one on ties)
        best = -1
        For tr = 0 To ntrks - 1
            If ptr(tr) < tend(tr) Then
                If best < 0 Then
                    best = tr
                ElseIf rawTick(ptr(tr)) < rawTick(ptr(best)) Then
                    best = tr
                End If
            End If
        Next
        If best < 0 Then Exit Do

        i = ptr(best)
        ptr(best) = i + 1
        tick = rawTick(i)
        msg = rawMsg(i)

        If (division And &H8000&) = 0 Then mspertick = tempo / (division * 1000#)
        t = t + (tick - lasttick) * mspertick
        lasttick = tick

        If (msg And &HFF000000) = MEVT_TEMPO Then
            tempo = msg And &HFFFFFF
        ElseIf msg <> MEVT_NOP Then
            AddEvent Int(t + 0.5), msg
        End If
        If t > endms Then endms = t
    Loop

    Erase rawTick
    Erase rawMsg
    AddEvent Int(endms + 0.5), MEVT_NOP          ' the song length, for looping
    ConvertMIDI = True
End Function

'
' Puts the next stretch of the song into midi buffer b (b = -1: -muslog
' only).  Returns False when there is nothing more to play.
'
Private Function FillChunk(ByVal b As Long) As Boolean
    Dim n As Long, o As Long, d As Long, span As Long, msg As Long, ch As Long

    If b >= 0 Then o = b * CHUNKEVENTS * 3

    Do While n < CHUNKEVENTS
        If evpos >= numev Then
            If Not musloop Then Exit Do
            evpos = loopstart
            lasttime = 0
        End If

        d = evTime(evpos) - lasttime
        If n > 0 And span + d > CHUNKMS Then Exit Do
        span = span + d

        msg = evMsg(evpos)
        If (msg And &HFF000000) = 0 Then
            If (msg And &HFFF0&) = &H7B0& Then   ' controller 7: channel volume
                ch = msg And 15
                chanvol(ch) = (msg \ 65536) And 127
                msg = (msg And &HFFFF&) Or ScaleVol(chanvol(ch)) * 65536
            End If
        End If

        streamtime = streamtime + d
        If muslog <> 0 Then Print #muslog, streamtime; Hex$(msg)
        If b >= 0 Then
            midibuf(o) = d                       ' MIDIEVENT: dwDeltaTime,
            midibuf(o + 1) = 0                   ' dwStreamID,
            midibuf(o + 2) = msg                 ' dwEvent
            o = o + 3
        End If

        n = n + 1
        lasttime = evTime(evpos)
        evpos = evpos + 1
    Loop

    If n = 0 Then Exit Function

    If b >= 0 Then
        mhdr(b).lpData = VarPtr(midibuf(b * CHUNKEVENTS * 3))
        mhdr(b).dwBufferLength = CHUNKEVENTS * 12
        mhdr(b).dwBytesRecorded = n * 12
        mhdr(b).dwFlags = 0
        mhdr(b).dwOffset = 0
        If midiOutPrepareHeader(hms, VarPtr(mhdr(b)), LenB(mhdr(b))) <> 0 Then Exit Function
        If midiStreamOut(hms, VarPtr(mhdr(b)), LenB(mhdr(b))) <> 0 Then
            midiOutUnprepareHeader hms, VarPtr(mhdr(b)), LenB(mhdr(b))
            Exit Function
        End If
        mqueued(b) = True
    End If

    FillChunk = True
End Function

Private Sub AllNotesOff()
    Dim ch As Long
    For ch = 0 To 15
        midiOutShortMsg hms, &HB0 Or ch Or 123 * 256&    ' all notes off
        midiOutShortMsg hms, &HB0 Or ch Or 120 * 256&    ' all sound off
    Next
End Sub

Public Sub I_InitMusic()
    Dim rc As Long, dev As Long, p As Long, ch As Long

    musvol = 15
    For ch = 0 To 15
        chanvol(ch) = 100
    Next
    ReDim evTime(0 To 1023)
    ReDim evMsg(0 To 1023)

    p = M_CheckParm("-muslog")
    If p > 0 And p < myargc - 1 Then
        muslog = FreeFile
        Open myargv(p + 1) For Output As #muslog
        musicon = True
        I_Printf "I_InitMusic: listing music events to " & myargv(p + 1)
        Exit Sub
    End If

    If M_CheckParm("-nosound") <> 0 Or M_CheckParm("-nomusic") <> 0 Then
        I_Printf "I_InitMusic: music off"
        Exit Sub
    End If

    dev = MIDI_MAPPER
    rc = midiStreamOpen(hms, dev, 1, 0, 0, 0)
    If rc <> 0 Then
        hms = 0
        I_Printf "I_InitMusic: could not open the MIDI device (error " & rc & ")"
        Exit Sub
    End If

    ' the event buffers stay at one address while the device owns them
    ReDim midibuf(0 To NUMMIDIBUFS * CHUNKEVENTS * 3 - 1)
    musicon = True
    I_Printf "I_InitMusic: MIDI music ready"
End Sub

Public Sub I_ShutdownMusic()
    Dim mmt(0 To 2) As Long

    If hms <> 0 Then
        If musplaying Then
            mmt(0) = TIME_TICKS
            midiStreamPosition hms, VarPtr(mmt(0)), 12
            I_Printf "I_ShutdownMusic: song position " & mmt(1) & " ms"
        End If
        I_StopSong 0
        midiStreamClose hms
        hms = 0
    End If
    If muslog <> 0 Then
        Close #muslog
        muslog = 0
    End If
    musicon = False
End Sub

Public Sub I_SetMusicVolume(ByVal volume As Long)
    Dim ch As Long

    ' Internal state variable (0..15; S_SetMusicVolume first passes 127).
    musvol = volume
    If musvol < 0 Then musvol = 0
    If musvol > 15 Then musvol = 15

    ' Now set volume on output device.
    If muslog <> 0 Then Print #muslog, "volume"; musvol
    If hms <> 0 And musplaying Then
        For ch = 0 To 15
            midiOutShortMsg hms, &HB0 Or ch Or 7 * 256& Or ScaleVol(chanvol(ch)) * 65536
        Next
    End If
End Sub

Public Sub I_PauseSong(ByVal handle As Long)
    If muslog <> 0 Then Print #muslog, "pause"
    If hms <> 0 And musplaying Then
        midiStreamPause hms
        AllNotesOff
    End If
End Sub

Public Sub I_ResumeSong(ByVal handle As Long)
    If muslog <> 0 Then Print #muslog, "resume"
    If hms <> 0 And musplaying Then midiStreamRestart hms
End Sub

Public Function I_RegisterSong(ByVal lump As Long) As Long
    Dim p As Long, size As Long

    songok = False
    numev = 0
    I_RegisterSong = 1
    If Not musicon Then Exit Function

    p = W_CacheLumpNum(lump)
    size = W_LumpLength(lump)
    AddPrelude
    If size > 16 Then
        If IsTag(p, "MUS" & Chr$(26)) Then
            songok = ConvertMUS(p, size)
        ElseIf IsTag(p, "MThd") Then
            songok = ConvertMIDI(p, size)
        End If
    End If

    ' a song needs a length, or looping it would never get anywhere
    If songok Then songok = (evTime(numev - 1) > 0)

    If Not songok Then
        I_Printf "I_RegisterSong: " & W_LumpName(lump) & " is not a MUS or MIDI song"
    ElseIf muslog <> 0 Then
        Print #muslog, "song "; W_LumpName(lump); numev - loopstart; "events"; evTime(numev - 1); "ms"
    End If
End Function

Public Sub I_PlaySong(ByVal handle As Long, ByVal looping As Boolean)
    Dim b As Long
    Dim prop(0 To 1) As Long

    If Not songok Then Exit Sub

    musloop = looping
    evpos = 0
    lasttime = 0
    streamtime = 0
    musplaying = True

    If hms <> 0 Then
        ' 1000 ticks per quarter note at 1000000 microseconds per quarter
        prop(0) = 8
        prop(1) = 1000
        midiStreamProperty hms, VarPtr(prop(0)), MIDIPROP_SET Or MIDIPROP_TIMEDIV
        prop(1) = 1000000
        midiStreamProperty hms, VarPtr(prop(0)), MIDIPROP_SET Or MIDIPROP_TEMPO
        For b = 0 To NUMMIDIBUFS - 1
            If Not FillChunk(b) Then Exit For
        Next
        midiStreamRestart hms
    ElseIf muslog <> 0 Then
        If looping Then Print #muslog, "play looping" Else Print #muslog, "play once"
        musstart = I_GetTimeMS()
        I_UpdateMusic
    End If
End Sub

Public Sub I_StopSong(ByVal handle As Long)
    Dim b As Long, tries As Long

    If muslog <> 0 And musplaying Then Print #muslog, "stop"
    musplaying = False
    If hms = 0 Then Exit Sub

    ' stop returns all queued buffers and turns the notes off
    midiStreamStop hms
    For b = 0 To NUMMIDIBUFS - 1
        If mqueued(b) Then
            tries = 0
            Do While midiOutUnprepareHeader(hms, VarPtr(mhdr(b)), LenB(mhdr(b))) = MIDIERR_STILLPLAYING
                tries = tries + 1
                If tries > 100 Then Exit Do
                Sleep 1
            Loop
            mqueued(b) = False
        End If
    Next
End Sub

Public Sub I_UnRegisterSong(ByVal handle As Long)
    songok = False
    numev = 0
End Sub

' Called every frame: refills the buffers the MIDI stream has played.
Public Sub I_UpdateMusic()
    Dim b As Long

    If Not musplaying Then Exit Sub

    If hms <> 0 Then
        For b = 0 To NUMMIDIBUFS - 1
            If mqueued(b) Then
                If (mhdr(b).dwFlags And MHDR_DONE) <> 0 Then
                    midiOutUnprepareHeader hms, VarPtr(mhdr(b)), LenB(mhdr(b))
                    mqueued(b) = False
                End If
            End If
        Next
        For b = 0 To NUMMIDIBUFS - 1
            If Not mqueued(b) Then
                If Not FillChunk(b) Then Exit For
            End If
        Next
    ElseIf muslog <> 0 Then
        Do While streamtime < I_GetTimeMS() - musstart + MUSLEADMS
            If Not FillChunk(-1) Then Exit Do
        Loop
    End If
End Sub
