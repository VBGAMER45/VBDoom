Attribute VB_Name = "f_finale"
' f_finale.bas - game completion, final screen animation (f_finale.c).
'
' The finale texts (E1TEXT ... T6TEXT) are the dstrings variables.
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' Stage of animation:
'  0 = text, 1 = art screen, 2 = character cast
Private finalestage As Long
Private finalecount As Long

Private Const TEXTSPEED As Long = 3
Private Const TEXTWAIT As Long = 250

Private finaletext As String
Private finaleflat As String

' Final DOOM 2 animation: casting by id Software, in order of appearance
' (C: castinfo_t castorder[], ended by a NULL name)
Private castname(0 To 16) As String
Private casttype(0 To 16) As Long
Private Const NUMCAST As Long = 17
Private castInit As Boolean

Private castnum As Long
Private casttics As Long
Private caststate As Long                        ' state number (C: state_t*)
Private castdeath As Boolean
Private castframes As Long
Private castonmelee As Long
Private castattacking As Boolean

' F_BunnyScroll static
Private laststage As Long

Private Sub InitCast()
    If castInit Then Exit Sub
    castInit = True
    castname(0) = CC_ZOMBIE: casttype(0) = MT_POSSESSED
    castname(1) = CC_SHOTGUN: casttype(1) = MT_SHOTGUY
    castname(2) = CC_HEAVY: casttype(2) = MT_CHAINGUY
    castname(3) = CC_IMP: casttype(3) = MT_TROOP
    castname(4) = CC_DEMON: casttype(4) = MT_SERGEANT
    castname(5) = CC_LOST: casttype(5) = MT_SKULL
    castname(6) = CC_CACO: casttype(6) = MT_HEAD
    castname(7) = CC_HELL: casttype(7) = MT_KNIGHT
    castname(8) = CC_BARON: casttype(8) = MT_BRUISER
    castname(9) = CC_ARACH: casttype(9) = MT_BABY
    castname(10) = CC_PAIN: casttype(10) = MT_PAIN
    castname(11) = CC_REVEN: casttype(11) = MT_UNDEAD
    castname(12) = CC_MANCU: casttype(12) = MT_FATSO
    castname(13) = CC_ARCH: casttype(13) = MT_VILE
    castname(14) = CC_SPIDER: casttype(14) = MT_SPIDER
    castname(15) = CC_CYBER: casttype(15) = MT_CYBORG
    castname(16) = CC_HERO: casttype(16) = MT_PLAYER
End Sub

' F_StartFinale
Public Sub F_StartFinale()
    gameaction = ga_nothing
    gamestate = GS_FINALE
    viewactive = False
    automapactive = False

    ' Okay - IWAD dependend stuff.
    ' This has been changed severly, and
    '  some stuff might have changed in the process.
    Select Case gamemode
        ' DOOM 1 - E1, E3 or E4, but each nine missions
        Case shareware, registered, retail
            S_ChangeMusic mus_victor, True

            Select Case gameepisode
                Case 1
                    finaleflat = "FLOOR4_8"
                    finaletext = E1TEXT
                Case 2
                    finaleflat = "SFLR6_1"
                    finaletext = E2TEXT
                Case 3
                    finaleflat = "MFLR8_4"
                    finaletext = E3TEXT
                Case 4
                    finaleflat = "MFLR8_3"
                    finaletext = E4TEXT
                Case Else
                    ' Ouch.
            End Select

        ' DOOM II and missions packs with E1, M34
        Case commercial
            S_ChangeMusic mus_read_m, True
            Select Case gamemap
                Case 6
                    finaleflat = "SLIME16"
                    finaletext = C1TEXT
                Case 11
                    finaleflat = "RROCK14"
                    finaletext = C2TEXT
                Case 20
                    finaleflat = "RROCK07"
                    finaletext = C3TEXT
                Case 30
                    finaleflat = "RROCK17"
                    finaletext = C4TEXT
                Case 15
                    finaleflat = "RROCK13"
                    finaletext = C5TEXT
                Case 31
                    finaleflat = "RROCK19"
                    finaletext = C6TEXT
                Case Else
                    ' Ouch.
            End Select

        ' Indeterminate.
        Case Else
            S_ChangeMusic mus_read_m, True
            finaleflat = "F_SKY1"                ' Not used anywhere else.
            finaletext = C1TEXT                  ' FIXME - other text, music?
    End Select

    finalestage = 0
    finalecount = 0
End Sub

Public Function F_Responder(ev As event_t) As Boolean
    If finalestage = 2 Then
        F_Responder = F_CastResponder(ev)
        Exit Function
    End If

    F_Responder = False
End Function

' F_Ticker
Public Sub F_Ticker()
    Dim i As Long

    ' check for skipping
    If gamemode = commercial And finalecount > 50 Then
        ' go on to the next level
        For i = 0 To MAXPLAYERS - 1
            If players(i).cmd.buttons <> 0 Then Exit For
        Next

        If i < MAXPLAYERS Then
            If gamemap = 30 Then
                F_StartCast
            Else
                gameaction = ga_worlddone
            End If
        End If
    End If

    ' advance animation
    finalecount = finalecount + 1

    If finalestage = 2 Then
        F_CastTicker
        Exit Sub
    End If

    If gamemode = commercial Then Exit Sub

    If finalestage = 0 And finalecount > Len(finaletext) * TEXTSPEED + TEXTWAIT Then
        finalecount = 0
        finalestage = 1
        wipegamestate = -1                       ' force a wipe
        If gameepisode = 3 Then S_StartMusic mus_bunny
    End If
End Sub

' F_TextWrite
Private Sub F_TextWrite()
    Dim src As Long, dest As Long, x As Long, y As Long, w As Long
    Dim count As Long, i As Long, c As Long, cx As Long, cy As Long

    ' erase the entire screen to a tiled background
    src = W_CacheLumpName(finaleflat)
    dest = 0                                     ' screens[0]

    For y = 0 To SCREENHEIGHT - 1
        For x = 0 To SCREENWIDTH \ 64 - 1
            CopyMemory screens(dest), wadbuf(src + (y And 63) * 64), 64
            dest = dest + 64
        Next
        If (SCREENWIDTH And 63) <> 0 Then
            CopyMemory screens(dest), wadbuf(src + (y And 63) * 64), SCREENWIDTH And 63
            dest = dest + (SCREENWIDTH And 63)
        End If
    Next
    V_MarkRect 0, 0, SCREENWIDTH, SCREENHEIGHT

    ' draw some of the text onto the screen
    cx = 10
    cy = 10
    i = 1

    count = (finalecount - 10) \ TEXTSPEED
    If count < 0 Then count = 0
    Do While count > 0
        If i > Len(finaletext) Then Exit Do
        c = Asc(Mid$(finaletext, i, 1))
        i = i + 1
        If c = 10 Then                           ' '\n'
            cx = 10
            cy = cy + 11
        Else
            ' (the C tested c > HU_FONTSIZE, one past the font)
            c = toupper(c) - HU_FONTSTART
            If c < 0 Or c >= HU_FONTSIZE Then
                cx = cx + 4
            Else
                w = PatchWidth(hu_font(c))
                If cx + w > SCREENWIDTH Then Exit Do
                V_DrawPatch cx, cy, 0, hu_font(c)
                cx = cx + w
            End If
        End If
        count = count - 1
    Loop
End Sub

' F_StartCast
Private Sub F_StartCast()
    InitCast
    wipegamestate = -1                           ' force a screen wipe
    castnum = 0
    caststate = mobjinfo(casttype(castnum)).seestate
    casttics = states(caststate).tics
    castdeath = False
    finalestage = 2
    castframes = 0
    castonmelee = 0
    castattacking = False
    S_ChangeMusic mus_evil, True
End Sub

' F_CastTicker
Private Sub F_CastTicker()
    Dim st As Long, sfx As Long, stopattack As Boolean

    casttics = casttics - 1
    If casttics > 0 Then Exit Sub                ' not time to change state yet

    If states(caststate).tics = -1 Or states(caststate).nextstate = S_NULL Then
        ' switch from deathstate to next monster
        castnum = castnum + 1
        castdeath = False
        If castnum = NUMCAST Then castnum = 0
        If mobjinfo(casttype(castnum)).seesound <> 0 Then S_StartSound 0, mobjinfo(casttype(castnum)).seesound
        caststate = mobjinfo(casttype(castnum)).seestate
        castframes = 0
    Else
        ' just advance to next state in animation
        If caststate = S_PLAY_ATK1 Then
            stopattack = True                    ' Oh, gross hack! (C: goto stopattack)
        Else
            st = states(caststate).nextstate
            caststate = st
            castframes = castframes + 1

            ' sound hacks....
            Select Case st
                Case S_PLAY_ATK1: sfx = sfx_dshtgn
                Case S_POSS_ATK2: sfx = sfx_pistol
                Case S_SPOS_ATK2: sfx = sfx_shotgn
                Case S_VILE_ATK2: sfx = sfx_vilatk
                Case S_SKEL_FIST2: sfx = sfx_skeswg
                Case S_SKEL_FIST4: sfx = sfx_skepch
                Case S_SKEL_MISS2: sfx = sfx_skeatk
                Case S_FATT_ATK8, S_FATT_ATK5, S_FATT_ATK2: sfx = sfx_firsht
                Case S_CPOS_ATK2, S_CPOS_ATK3, S_CPOS_ATK4: sfx = sfx_shotgn
                Case S_TROO_ATK3: sfx = sfx_claw
                Case S_SARG_ATK2: sfx = sfx_sgtatk
                Case S_BOSS_ATK2, S_BOS2_ATK2, S_HEAD_ATK2: sfx = sfx_firsht
                Case S_SKULL_ATK2: sfx = sfx_sklatk
                Case S_SPID_ATK2, S_SPID_ATK3: sfx = sfx_shotgn
                Case S_BSPI_ATK2: sfx = sfx_plasma
                Case S_CYBER_ATK2, S_CYBER_ATK4, S_CYBER_ATK6: sfx = sfx_rlaunc
                Case S_PAIN_ATK3: sfx = sfx_sklatk
                Case Else: sfx = 0
            End Select

            If sfx <> 0 Then S_StartSound 0, sfx
        End If
    End If

    If Not stopattack Then
        If castframes = 12 Then
            ' go into attack frame
            castattacking = True
            If castonmelee <> 0 Then
                caststate = mobjinfo(casttype(castnum)).meleestate
            Else
                caststate = mobjinfo(casttype(castnum)).missilestate
            End If
            castonmelee = castonmelee Xor 1
            If caststate = S_NULL Then
                If castonmelee <> 0 Then
                    caststate = mobjinfo(casttype(castnum)).meleestate
                Else
                    caststate = mobjinfo(casttype(castnum)).missilestate
                End If
            End If
        End If

        If castattacking Then
            If castframes = 24 Or caststate = mobjinfo(casttype(castnum)).seestate Then stopattack = True
        End If
    End If

    If stopattack Then
        castattacking = False
        castframes = 0
        caststate = mobjinfo(casttype(castnum)).seestate
    End If

    casttics = states(caststate).tics
    If casttics = -1 Then casttics = 15
End Sub

' F_CastResponder
Private Function F_CastResponder(ev As event_t) As Boolean
    If ev.typ <> ev_keydown Then
        F_CastResponder = False
        Exit Function
    End If

    If castdeath Then
        F_CastResponder = True                   ' already in dying frames
        Exit Function
    End If

    ' go into death frame
    castdeath = True
    caststate = mobjinfo(casttype(castnum)).deathstate
    casttics = states(caststate).tics
    castframes = 0
    castattacking = False
    If mobjinfo(casttype(castnum)).deathsound <> 0 Then S_StartSound 0, mobjinfo(casttype(castnum)).deathsound

    F_CastResponder = True
End Function

Private Sub F_CastPrint(ByVal text As String)
    Dim i As Long, c As Long, cx As Long, w As Long, width As Long

    ' find width
    width = 0
    For i = 1 To Len(text)
        c = toupper(Asc(Mid$(text, i, 1))) - HU_FONTSTART
        If c < 0 Or c >= HU_FONTSIZE Then
            width = width + 4
        Else
            w = PatchWidth(hu_font(c))
            width = width + w
        End If
    Next

    ' draw it
    cx = 160 - width \ 2
    For i = 1 To Len(text)
        c = toupper(Asc(Mid$(text, i, 1))) - HU_FONTSTART
        If c < 0 Or c >= HU_FONTSIZE Then
            cx = cx + 4
        Else
            w = PatchWidth(hu_font(c))
            V_DrawPatch cx, 180, 0, hu_font(c)
            cx = cx + w
        End If
    Next
End Sub

' F_CastDrawer
Private Sub F_CastDrawer()
    Dim fr As Long, lump As Long, flip As Boolean

    ' erase the entire screen to a background
    V_DrawPatch 0, 0, 0, W_GetNumForName("BOSSBACK")

    F_CastPrint castname(castnum)

    ' draw the current frame in the middle of the screen
    fr = spr_firstframe(states(caststate).sprite) + (states(caststate).frame And FF_FRAMEMASK)
    lump = sf_lump(fr * 8)
    flip = (sf_flip(fr * 8) <> 0)

    If flip Then
        V_DrawPatchFlipped 160, 170, 0, lump + firstspritelump
    Else
        V_DrawPatch 160, 170, 0, lump + firstspritelump
    End If
End Sub

' F_DrawPatchCol - draws column col of a patch, without its offsets
Private Sub F_DrawPatchCol(ByVal x As Long, ByVal patch As Long, ByVal col As Long)
    Dim p As Long, column As Long, source As Long, dest As Long, count As Long

    p = lumpinfo(patch).position
    column = p + RdLong(p + 8 + col * 4)

    ' step through the posts in a column
    Do While wadbuf(column) <> &HFF
        source = column + 3
        dest = x + wadbuf(column) * SCREENWIDTH
        count = wadbuf(column + 1)

        Do While count > 0
            screens(dest) = wadbuf(source)
            source = source + 1
            dest = dest + SCREENWIDTH
            count = count - 1
        Loop
        column = column + wadbuf(column + 1) + 4
    Loop
End Sub

' F_BunnyScroll
Private Sub F_BunnyScroll()
    Dim scrolled As Long, x As Long, p1 As Long, p2 As Long, stage As Long

    p1 = W_GetNumForName("PFUB2")
    p2 = W_GetNumForName("PFUB1")

    V_MarkRect 0, 0, SCREENWIDTH, SCREENHEIGHT

    scrolled = 320 - (finalecount - 230) \ 2
    If scrolled > 320 Then scrolled = 320
    If scrolled < 0 Then scrolled = 0

    For x = 0 To SCREENWIDTH - 1
        If x + scrolled < 320 Then
            F_DrawPatchCol x, p1, x + scrolled
        Else
            F_DrawPatchCol x, p2, x + scrolled - 320
        End If
    Next

    If finalecount < 1130 Then Exit Sub
    If finalecount < 1180 Then
        V_DrawPatch (SCREENWIDTH - 13 * 8) \ 2, (SCREENHEIGHT - 8 * 8) \ 2, 0, W_GetNumForName("END0")
        laststage = 0
        Exit Sub
    End If

    stage = (finalecount - 1180) \ 5
    If stage > 6 Then stage = 6
    If stage > laststage Then
        S_StartSound 0, sfx_pistol
        laststage = stage
    End If

    V_DrawPatch (SCREENWIDTH - 13 * 8) \ 2, (SCREENHEIGHT - 8 * 8) \ 2, 0, W_GetNumForName("END" & stage)
End Sub

' F_Drawer
Public Sub F_Drawer()
    If finalestage = 2 Then
        F_CastDrawer
        Exit Sub
    End If

    If finalestage = 0 Then
        F_TextWrite
    Else
        Select Case gameepisode
            Case 1
                If gamemode = retail Then
                    V_DrawPatch 0, 0, 0, W_GetNumForName("CREDIT")
                Else
                    V_DrawPatch 0, 0, 0, W_GetNumForName("HELP2")
                End If
            Case 2
                V_DrawPatch 0, 0, 0, W_GetNumForName("VICTORY2")
            Case 3
                F_BunnyScroll
            Case 4
                V_DrawPatch 0, 0, 0, W_GetNumForName("ENDPIC")
        End Select
    End If
End Sub
