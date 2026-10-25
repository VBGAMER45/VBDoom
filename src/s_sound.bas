Attribute VB_Name = "s_sound"
' s_sound.c - sound effect channels, positional sound and music selection
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
'
' Origins (C: void* / mobj_t*): 0 = none, > 0 = mobj slot, < 0 = sector sound
' origin -(sector + 1) (SectorSoundOrg).
'
' Volume: the menu and default.cfg keep snd_SfxVolume / snd_MusicVolume at
' 0..15.  The DOS version passed snd_SfxVolume * 8 to S_Init and
' S_SetSfxVolume; linuxdoom commented the "* 8" out, which made effects eight
' times quieter and broke the distance falloff on map 8.  Here the 0..127
' value lives in sfxvol and is used wherever the C read snd_SfxVolume.
Option Explicit

Private Const S_MAX_VOLUME As Long = 127

' when to clip out sounds
' Does not fit the large outdoor areas.
Private Const S_CLIPPING_DIST As Long = 1200& * &H10000

' Distance tp origin when sounds should be maxed out.
' This should relate to movement clipping resolution
' (see BLOCKMAP handling).
' Originally: (200*0x10000).
Private Const S_CLOSE_DIST As Long = 160& * &H10000

Private Const S_ATTENUATOR As Long = (S_CLIPPING_DIST - S_CLOSE_DIST) \ &H10000

Private Const NORM_PITCH As Long = 128
Private Const NORM_PRIORITY As Long = 64
Private Const NORM_SEP As Long = 128

Private Const S_STEREO_SWING As Long = 96& * &H10000

' the set of channels available (C: channel_t*)
Private chSfx() As Long                          ' C: sfxinfo (sfx number, 0 = channel free)
Private chOrigin() As Long                       ' origin of sound
Private chHandle() As Long                       ' handle of the sound being played

Private sfxvol As Long                           ' snd_SfxVolume * 8 (see above)

' whether songs are mus_paused
Private mus_paused As Boolean

' music currently being played (C: musicinfo_t*; music number, 0 = none)
Private mus_playing As Long

Private nextcleanup As Long

' episode 4 songs (C: spmus, local to S_Start)
Private spmus(0 To 8) As Long

'
' Initializes sound stuff, including volume
' Sets channels, SFX and music volume,
'  allocates channel buffer, sets S_sfx lookup.
'
Public Sub S_Init(ByVal sfxVolume As Long, ByVal musicVolume As Long)
    Dim i As Long

    I_Printf "S_Init: default sfx volume " & sfxVolume

    ' Whatever these did with DMX, these are rather dummies now.
    I_SetChannels

    S_SetSfxVolume sfxVolume
    S_SetMusicVolume musicVolume

    ' Allocating the internal channels for mixing
    ' (the maximum numer of sounds rendered
    ' simultaneously).  One spare entry, so that 0 channels works.
    If numChannels < 0 Then numChannels = 0
    ReDim chSfx(0 To numChannels)
    ReDim chOrigin(0 To numChannels)
    ReDim chHandle(0 To numChannels)

    ' Free all channels for use
    For i = 0 To numChannels - 1
        chSfx(i) = 0
    Next

    ' no sounds are playing, and they are not mus_paused
    mus_paused = False

    ' Note that sounds have not been cached (yet).
    For i = 1 To NUMSFX - 1
        S_sfx(i).lumpnum = -1
        S_sfx(i).usefulness = -1
    Next

    ' Song - Who? - Where?
    spmus(0) = mus_e3m4                          ' American     e4m1
    spmus(1) = mus_e3m2                          ' Romero       e4m2
    spmus(2) = mus_e3m3                          ' Shawn        e4m3
    spmus(3) = mus_e1m5                          ' American     e4m4
    spmus(4) = mus_e2m7                          ' Tim          e4m5
    spmus(5) = mus_e2m4                          ' Romero       e4m6
    spmus(6) = mus_e2m6                          ' J.Anderson   e4m7 CHIRON.WAD
    spmus(7) = mus_e2m5                          ' Shawn        e4m8
    spmus(8) = mus_e1m9                          ' Tim          e4m9
End Sub

'
' Per level startup code.
' Kills playing sounds at start of level,
'  determines music if any, changes music.
'
Public Sub S_Start()
    Dim cnum As Long, mnum As Long

    ' kill all playing sounds at start of level
    '  (trust me - a good idea)
    For cnum = 0 To numChannels - 1
        If chSfx(cnum) <> 0 Then S_StopChannel cnum
    Next

    ' start new music for the level
    mus_paused = False

    If gamemode = commercial Then
        mnum = mus_runnin + gamemap - 1
    Else
        If gameepisode < 4 Then
            mnum = mus_e1m1 + (gameepisode - 1) * 9 + gamemap - 1
        Else
            mnum = spmus(gamemap - 1)
        End If
    End If

    S_ChangeMusic mnum, True

    nextcleanup = 15
End Sub

Public Sub S_StartSoundAtVolume(ByVal origin As Long, ByVal sfx_id As Long, ByVal volume As Long)
    Dim rc As Boolean, sep As Long, pitch As Long, priority As Long, cnum As Long
    Dim listener As Long, ox As Long, oy As Long

    ' check for bogus sound # (the C tested > NUMSFX)
    If sfx_id < 1 Or sfx_id >= NUMSFX Then I_Error "Bad sfx #: " & sfx_id

    ' Initialize sound parameters
    If S_sfx(sfx_id).link >= 0 Then
        pitch = S_sfx(sfx_id).pitch
        priority = S_sfx(sfx_id).priority
        volume = volume + S_sfx(sfx_id).volume

        If volume < 1 Then Exit Sub

        If volume > sfxvol Then volume = sfxvol
    Else
        pitch = NORM_PITCH
        priority = NORM_PRIORITY
    End If

    ' Check to see if it is audible,
    '  and if not, modify the params
    listener = players(consoleplayer).mo
    If origin <> 0 And origin <> listener Then
        rc = S_AdjustSoundParams(listener, origin, volume, sep, pitch)

        If listener <> 0 Then
            OriginXY origin, ox, oy
            If ox = mobjs(listener).x Then
                If oy = mobjs(listener).y Then sep = NORM_SEP
            End If
        End If

        If Not rc Then Exit Sub
    Else
        sep = NORM_SEP
    End If

    ' hacks to vary the sfx pitches
    If sfx_id >= sfx_sawup And sfx_id <= sfx_sawhit Then
        pitch = pitch + 8 - (M_Random() And 15)

        If pitch < 0 Then
            pitch = 0
        ElseIf pitch > 255 Then
            pitch = 255
        End If
    ElseIf sfx_id <> sfx_itemup And sfx_id <> sfx_tink Then
        pitch = pitch + 16 - (M_Random() And 31)

        If pitch < 0 Then
            pitch = 0
        ElseIf pitch > 255 Then
            pitch = 255
        End If
    End If

    ' kill old sound
    S_StopSound origin

    ' try to find a channel
    cnum = S_getChannel(origin, sfx_id)

    If cnum < 0 Then Exit Sub

    ' get lumpnum if necessary
    If S_sfx(sfx_id).lumpnum < 0 Then S_sfx(sfx_id).lumpnum = I_GetSfxLumpNum(sfx_id)

    ' increase the usefulness
    If S_sfx(sfx_id).usefulness < 0 Then
        S_sfx(sfx_id).usefulness = 1
    Else
        S_sfx(sfx_id).usefulness = S_sfx(sfx_id).usefulness + 1
    End If

    ' Assigns the handle to one of the channels in the
    '  mix/output buffer.
    chHandle(cnum) = I_StartSound(sfx_id, volume, sep, pitch, priority)
End Sub

Public Sub S_StartSound(ByVal origin As Long, ByVal sfx_id As Long)
    S_StartSoundAtVolume origin, sfx_id, sfxvol
End Sub

Public Sub S_StopSound(ByVal origin As Long)
    Dim cnum As Long

    For cnum = 0 To numChannels - 1
        If chSfx(cnum) <> 0 Then
            If chOrigin(cnum) = origin Then
                S_StopChannel cnum
                Exit For
            End If
        End If
    Next
End Sub

'
' Stop and resume music, during game PAUSE.
'
Public Sub S_PauseSound()
    If mus_playing <> 0 And Not mus_paused Then
        I_PauseSong S_music(mus_playing).handle
        mus_paused = True
    End If
End Sub

Public Sub S_ResumeSound()
    If mus_playing <> 0 And mus_paused Then
        I_ResumeSong S_music(mus_playing).handle
        mus_paused = False
    End If
End Sub

'
' Updates music & sounds
'
Public Sub S_UpdateSounds(ByVal listener As Long)
    Dim audible As Boolean, cnum As Long, volume As Long, sep As Long, pitch As Long
    Dim sfx As Long

    ' (the C's cleanup of unused sound data was commented out)

    For cnum = 0 To numChannels - 1
        sfx = chSfx(cnum)

        If sfx <> 0 Then
            If I_SoundIsPlaying(chHandle(cnum)) Then
                ' initialize parameters
                volume = sfxvol
                pitch = NORM_PITCH
                sep = NORM_SEP

                If S_sfx(sfx).link >= 0 Then
                    pitch = S_sfx(sfx).pitch
                    volume = volume + S_sfx(sfx).volume
                    If volume < 1 Then
                        S_StopChannel cnum
                        GoTo nextchannel
                    ElseIf volume > sfxvol Then
                        volume = sfxvol
                    End If
                End If

                ' check non-local sounds for distance clipping
                '  or modify their params
                If chOrigin(cnum) <> 0 And listener <> chOrigin(cnum) Then
                    audible = S_AdjustSoundParams(listener, chOrigin(cnum), volume, sep, pitch)

                    If Not audible Then
                        S_StopChannel cnum
                    Else
                        I_UpdateSoundParams chHandle(cnum), volume, sep, pitch
                    End If
                End If
            Else
                ' if channel is allocated but sound has stopped,
                '  free it
                S_StopChannel cnum
            End If
        End If
nextchannel:
    Next
    ' (the C's "kill music if it is a single-play && finished" was commented out)
End Sub

Public Sub S_SetMusicVolume(ByVal volume As Long)
    If volume < 0 Or volume > 127 Then
        I_Error "Attempt to set music volume at " & volume
    End If

    I_SetMusicVolume 127
    I_SetMusicVolume volume
    snd_MusicVolume = volume
End Sub

Public Sub S_SetSfxVolume(ByVal volume As Long)
    If volume < 0 Or volume > 127 Then I_Error "Attempt to set sfx volume at " & volume

    snd_SfxVolume = volume
    sfxvol = volume * 8
    If sfxvol > S_MAX_VOLUME Then sfxvol = S_MAX_VOLUME
End Sub

'
' Starts some music with the music id found in sounds.h.
'
Public Sub S_StartMusic(ByVal music_id As Long)
    S_ChangeMusic music_id, False
End Sub

Public Sub S_ChangeMusic(ByVal music_id As Long, ByVal looping As Boolean)
    If music_id <= mus_None Or music_id >= NUMMUSIC Then
        I_Error "Bad music number " & music_id
    End If

    If mus_playing = music_id Then Exit Sub

    ' shutdown old music
    S_StopMusic

    ' get lumpnum if neccessary (the C bombed out when the lump was
    ' missing; a WAD without the song simply plays none)
    If S_music(music_id).lumpnum = 0 Then
        S_music(music_id).lumpnum = W_CheckNumForName("d_" & S_music(music_id).name)
    End If
    If S_music(music_id).lumpnum < 0 Then Exit Sub

    ' load & register it
    S_music(music_id).handle = I_RegisterSong(S_music(music_id).lumpnum)

    ' play it
    I_PlaySong S_music(music_id).handle, looping

    mus_playing = music_id
End Sub

Public Sub S_StopMusic()
    If mus_playing <> 0 Then
        If mus_paused Then I_ResumeSong S_music(mus_playing).handle

        I_StopSong S_music(mus_playing).handle
        I_UnRegisterSong S_music(mus_playing).handle

        mus_playing = 0
    End If
End Sub

Private Sub S_StopChannel(ByVal cnum As Long)
    Dim sfx As Long

    sfx = chSfx(cnum)
    If sfx <> 0 Then
        ' stop the sound playing
        If I_SoundIsPlaying(chHandle(cnum)) Then I_StopSound chHandle(cnum)

        ' (the C then looked for other channels playing the same sound
        ' and did nothing with the result)

        ' degrade usefulness of sound data
        S_sfx(sfx).usefulness = S_sfx(sfx).usefulness - 1

        chSfx(cnum) = 0
    End If
End Sub

' Position of a sound origin (C: origin->x, origin->y; a sector's soundorg
' is a degenmobj_t with the same layout).
Private Sub OriginXY(ByVal origin As Long, x As Long, y As Long)
    If origin > 0 Then
        x = mobjs(origin).x
        y = mobjs(origin).y
    Else
        x = sectors(-origin - 1).soundorgx
        y = sectors(-origin - 1).soundorgy
    End If
End Sub

'
' Changes volume, stereo-separation, and pitch variables
'  from the norm of a sound effect to be played.
' If the sound is not audible, returns False.
' Otherwise, modifies parameters and returns True.
'
Private Function S_AdjustSoundParams(ByVal listener As Long, ByVal source As Long, _
        vol As Long, sep As Long, pitch As Long) As Boolean
    Dim approx_dist As Long, adx As Long, ady As Long, angle As Long
    Dim sx As Long, sy As Long

    ' no listener mobj (the C would dereference NULL): play it unchanged
    If listener = 0 Then
        sep = NORM_SEP
        S_AdjustSoundParams = (vol > 0)
        Exit Function
    End If

    ' calculate the distance to sound origin
    '  and clip it if necessary
    OriginXY source, sx, sy
    adx = Abs(mobjs(listener).x - sx)
    ady = Abs(mobjs(listener).y - sy)

    ' From _GG1_ p.428. Appox. eucledian distance fast.
    If adx < ady Then
        approx_dist = adx + ady - adx \ 2
    Else
        approx_dist = adx + ady - ady \ 2
    End If

    If gamemap <> 8 And approx_dist > S_CLIPPING_DIST Then Exit Function

    ' angle of source to listener
    angle = R_PointToAngle2(mobjs(listener).x, mobjs(listener).y, sx, sy)

    If UGT(angle, mobjs(listener).angle) Then
        angle = angle - mobjs(listener).angle
    Else
        angle = angle + (&HFFFFFFFF - mobjs(listener).angle)
    End If

    angle = FINEANGLE(angle)

    ' stereo separation
    sep = 128 - SAR(FixedMul(S_STEREO_SWING, finesine(angle)), FRACBITS)

    ' volume calculation
    If approx_dist < S_CLOSE_DIST Then
        vol = sfxvol
    ElseIf gamemap = 8 Then
        If approx_dist > S_CLIPPING_DIST Then approx_dist = S_CLIPPING_DIST

        vol = 15 + ((sfxvol - 15) * SAR(S_CLIPPING_DIST - approx_dist, FRACBITS)) \ S_ATTENUATOR
    Else
        ' distance effect
        vol = (sfxvol * SAR(S_CLIPPING_DIST - approx_dist, FRACBITS)) \ S_ATTENUATOR
    End If

    S_AdjustSoundParams = (vol > 0)
End Function

'
' S_getChannel :
'   If none available, return -1.  Otherwise channel #.
'
Private Function S_getChannel(ByVal origin As Long, ByVal sfx As Long) As Long
    ' channel number to use
    Dim cnum As Long

    ' Find an open channel
    For cnum = 0 To numChannels - 1
        If chSfx(cnum) = 0 Then
            Exit For
        ElseIf origin <> 0 And chOrigin(cnum) = origin Then
            S_StopChannel cnum
            Exit For
        End If
    Next

    ' None available
    If cnum = numChannels Then
        ' Look for lower priority
        For cnum = 0 To numChannels - 1
            If S_sfx(chSfx(cnum)).priority >= S_sfx(sfx).priority Then Exit For
        Next

        If cnum = numChannels Then
            ' No lower priority.  Sorry, Charlie.
            S_getChannel = -1
            Exit Function
        Else
            ' Otherwise, kick out lower priority.
            S_StopChannel cnum
        End If
    End If

    ' channel is decided to be cnum.
    chSfx(cnum) = sfx
    chOrigin(cnum) = origin

    S_getChannel = cnum
End Function
