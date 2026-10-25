Attribute VB_Name = "st_stuff"
' st_stuff.bas - status bar code.  Does the face/direction indicator
' animation and the palette indicators as well (red pain/berserk, bright
' pickup); also handles the cheat codes (st_stuff.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' Background and foreground screen numbers
Private Const BG As Long = 4
Private Const FG As Long = 0

' Palette indices.  For damage/bonus red-/gold-shifts
Private Const STARTREDPALS As Long = 1
Private Const STARTBONUSPALS As Long = 9
Private Const NUMREDPALS As Long = 8
Private Const NUMBONUSPALS As Long = 4
' Radiation suit, green shift.
Private Const RADIATIONPAL As Long = 13

' N/256*100% probability that the normal face state will change
Private Const ST_FACEPROBABILITY As Long = 96

' For Responder
Private Const ST_TOGGLECHAT As Long = KEY_ENTER

' Location of status bar
Private Const ST_X As Long = 0
Private Const ST_X2 As Long = 104

Private Const ST_FX As Long = 143
Private Const ST_FY As Long = 169

' Number of status faces.
Private Const ST_NUMPAINFACES As Long = 5
Private Const ST_NUMSTRAIGHTFACES As Long = 3
Private Const ST_NUMTURNFACES As Long = 2
Private Const ST_NUMSPECIALFACES As Long = 3

Private Const ST_FACESTRIDE As Long = ST_NUMSTRAIGHTFACES + ST_NUMTURNFACES + ST_NUMSPECIALFACES

Private Const ST_NUMEXTRAFACES As Long = 2

Private Const ST_NUMFACES As Long = ST_FACESTRIDE * ST_NUMPAINFACES + ST_NUMEXTRAFACES

Private Const ST_TURNOFFSET As Long = ST_NUMSTRAIGHTFACES
Private Const ST_OUCHOFFSET As Long = ST_TURNOFFSET + ST_NUMTURNFACES
Private Const ST_EVILGRINOFFSET As Long = ST_OUCHOFFSET + 1
Private Const ST_RAMPAGEOFFSET As Long = ST_EVILGRINOFFSET + 1
Private Const ST_GODFACE As Long = ST_NUMPAINFACES * ST_FACESTRIDE
Private Const ST_DEADFACE As Long = ST_GODFACE + 1

Private Const ST_FACESX As Long = 143
Private Const ST_FACESY As Long = 168

Private Const ST_EVILGRINCOUNT As Long = 2 * TICRATE
Private Const ST_STRAIGHTFACECOUNT As Long = TICRATE \ 2
Private Const ST_TURNCOUNT As Long = 1 * TICRATE
Private Const ST_OUCHCOUNT As Long = 1 * TICRATE
Private Const ST_RAMPAGEDELAY As Long = 2 * TICRATE

Private Const ST_MUCHPAIN As Long = 20

' Location and size of statistics, justified according to widget type.

' AMMO number pos.
Private Const ST_AMMOWIDTH As Long = 3
Private Const ST_AMMOX As Long = 44
Private Const ST_AMMOY As Long = 171

' HEALTH number pos.
Private Const ST_HEALTHWIDTH As Long = 3
Private Const ST_HEALTHX As Long = 90
Private Const ST_HEALTHY As Long = 171

' Weapon pos.
Private Const ST_ARMSX As Long = 111
Private Const ST_ARMSY As Long = 172
Private Const ST_ARMSBGX As Long = 104
Private Const ST_ARMSBGY As Long = 168
Private Const ST_ARMSXSPACE As Long = 12
Private Const ST_ARMSYSPACE As Long = 10

' Frags pos.
Private Const ST_FRAGSX As Long = 138
Private Const ST_FRAGSY As Long = 171
Private Const ST_FRAGSWIDTH As Long = 2

' ARMOR number pos.
Private Const ST_ARMORWIDTH As Long = 3
Private Const ST_ARMORX As Long = 221
Private Const ST_ARMORY As Long = 171

' Key icon positions.
Private Const ST_KEY0X As Long = 239
Private Const ST_KEY0Y As Long = 171
Private Const ST_KEY1X As Long = 239
Private Const ST_KEY1Y As Long = 181
Private Const ST_KEY2X As Long = 239
Private Const ST_KEY2Y As Long = 191

' Ammunition counter.
Private Const ST_AMMO0WIDTH As Long = 3
Private Const ST_AMMO0X As Long = 288
Private Const ST_AMMO0Y As Long = 173
Private Const ST_AMMO1X As Long = 288
Private Const ST_AMMO1Y As Long = 179
Private Const ST_AMMO2X As Long = 288
Private Const ST_AMMO2Y As Long = 191
Private Const ST_AMMO3X As Long = 288
Private Const ST_AMMO3Y As Long = 185

' Indicate maximum ammunition.  Only needed because backpack exists.
Private Const ST_MAXAMMO0WIDTH As Long = 3
Private Const ST_MAXAMMO0X As Long = 314
Private Const ST_MAXAMMO0Y As Long = 173
Private Const ST_MAXAMMO1X As Long = 314
Private Const ST_MAXAMMO1Y As Long = 179
Private Const ST_MAXAMMO2X As Long = 314
Private Const ST_MAXAMMO2Y As Long = 191
Private Const ST_MAXAMMO3X As Long = 314
Private Const ST_MAXAMMO3Y As Long = 185

Private Const ST_MSGWIDTH As Long = 52

' st_stateenum_t
Private Const AutomapState As Long = 0
Private Const FirstPersonState As Long = 1

' st_chatstateenum_t
Private Const StartChatState As Long = 0
Private Const WaitDestState As Long = 1
Private Const GetChatState As Long = 2

' main player in game (C: player_t* plyr)
Private plyr As Long

' ST_Start() has just been called
Private st_firsttime As Boolean

' used to execute ST_Init() only once
Private veryfirsttime As Long

' lump number for PLAYPAL
Private lu_palette As Long

' used for timing
Private st_clock As Long

' used for making messages go away
Private st_msgcounter As Long

' used when in chat
Private st_chatstate As Long

' whether in automap or first-person
Private st_gamestate As Long

' whether left-side main status bar is active
Private st_statusbaron As Boolean

' whether status bar chat is active
Private st_chat As Boolean

' value of st_chat before message popped up
Private st_oldchat As Boolean

' whether chat window has the cursor on
Private st_cursoron As Boolean

' !deathmatch
Private st_notdeathmatch As Boolean

' !deathmatch && st_statusbaron
Private st_armson As Boolean

' !deathmatch
Private st_fragson As Boolean

' main bar left
Private sbar As Long

' 0-9, tall numbers
Private tallnum(0 To 9) As Long

' tall % sign
Private tallpercent As Long

' 0-9, short, yellow (,different!) numbers
Private shortnum(0 To 9) As Long

' 3 key-cards, 3 skulls
Private keys(0 To 5) As Long                     ' NUMCARDS

' face status patches
Private faces(0 To 41) As Long                   ' ST_NUMFACES

' face background
Private faceback As Long

' main bar right
Private armsbg As Long

' weapon ownership patches (C: arms[6][2], here arms(i * 2 + j))
Private arms(0 To 11) As Long

' ready-weapon widget
Private w_ready As st_number_t

' in deathmatch only, summary of frags stats
Private w_frags As st_number_t

' health widget
Private w_health As st_percent_t

' arms background
Private w_armsbg As st_binicon_t

' weapon ownership widgets
Private w_arms(0 To 5) As st_multicon_t

' face status widget
Private w_faces As st_multicon_t

' keycard widgets
Private w_keyboxes(0 To 2) As st_multicon_t

' armor widget
Private w_armor As st_percent_t

' ammo widgets
Private w_ammo(0 To 3) As st_number_t

' max ammo widgets
Private w_maxammo(0 To 3) As st_number_t

' ammo type shown by w_ready, -1 = none (C: w_ready.num = &largeammo)
Private st_readyammo As Long

' number of frags so far in deathmatch
Private st_fragscount As Long

' used to use appopriately pained face
Private st_oldhealth As Long

' used for evil grin
Private oldweaponsowned(0 To 8) As Boolean       ' NUMWEAPONS

' count until face changes
Private st_facecount As Long

' current face index, used by w_faces
Private st_faceindex As Long

' holds key-type for each key box on bar
Private keyboxes(0 To 2) As Long

' a random number per tick
Private st_randomnumber As Long

Private st_palette As Long
Private st_stopped As Boolean

' ST_calcPainOffset statics
Private lastcalc As Long
Private oldhealth As Long

' ST_updateFaceWidget statics
Private lastattackdown As Long
Private priority As Long

' Massive bunches of cheat shit to keep it from being easy to figure them
' out.  Yeah, right...  (cht_MakeSeq scrambles the plain text.)
Private cheat_mus As cheatseq_t
Private cheat_god As cheatseq_t
Private cheat_ammo As cheatseq_t
Private cheat_ammonokey As cheatseq_t
Private cheat_noclip As cheatseq_t
Private cheat_commercial_noclip As cheatseq_t
Private cheat_powerup(0 To 6) As cheatseq_t
Private cheat_choppers As cheatseq_t
Private cheat_clev As cheatseq_t
Private cheat_mypos As cheatseq_t

Private Sub InitCheats()
    cht_MakeSeq cheat_mus, "idmus" & Chr$(1) & Chr$(0) & Chr$(0)
    cht_MakeSeq cheat_choppers, "idchoppers"
    cht_MakeSeq cheat_god, "iddqd"
    cht_MakeSeq cheat_ammo, "idkfa"
    cht_MakeSeq cheat_ammonokey, "idfa"
    ' Smashing Pumpkins Into Samml Piles Of Putried Debris.
    cht_MakeSeq cheat_noclip, "idspispopd"
    cht_MakeSeq cheat_commercial_noclip, "idclip"
    cht_MakeSeq cheat_powerup(0), "idbeholdv"
    cht_MakeSeq cheat_powerup(1), "idbeholds"
    cht_MakeSeq cheat_powerup(2), "idbeholdi"
    cht_MakeSeq cheat_powerup(3), "idbeholdr"
    cht_MakeSeq cheat_powerup(4), "idbeholda"
    cht_MakeSeq cheat_powerup(5), "idbeholdl"
    cht_MakeSeq cheat_powerup(6), "idbehold"
    cht_MakeSeq cheat_clev, "idclev" & Chr$(1) & Chr$(0) & Chr$(0)
    ' my position cheat
    cht_MakeSeq cheat_mypos, "idmypos"
End Sub

' STATUS BAR CODE

Private Sub ST_refreshBackground()
    If st_statusbaron Then
        V_DrawPatch ST_X, 0, BG, sbar

        If netgame Then V_DrawPatch ST_FX, 0, BG, faceback

        V_CopyRect ST_X, 0, BG, ST_WIDTH, ST_HEIGHT, ST_X, ST_Y, FG
    End If
End Sub

' Respond to keyboard input events, intercept cheats.
Public Function ST_Responder(ev As event_t) As Boolean
    Dim i As Long, buf As String, musnum As Long, epsd As Long, map As Long
    Dim hit As Boolean, mo As Long

    ' Filter automap on/off.
    If ev.typ = ev_keyup And (ev.data1 And &HFFFF0000) = AM_MSGHEADER Then
        Select Case ev.data1
            Case AM_MSGENTERED
                st_gamestate = AutomapState
                st_firsttime = True
            Case AM_MSGEXITED
                st_gamestate = FirstPersonState
        End Select

    ' if a user keypress...
    ElseIf ev.typ = ev_keydown Then
        If Not netgame Then
            ' b. - enabled for more debug fun.
            ' if (gameskill != sk_nightmare) {

            ' 'dqd' cheat for toggleable god mode
            If cht_CheckCheat(cheat_god, ev.data1) Then
                players(plyr).cheats = players(plyr).cheats Xor CF_GODMODE
                If (players(plyr).cheats And CF_GODMODE) <> 0 Then
                    If players(plyr).mo <> 0 Then mobjs(players(plyr).mo).health = 100

                    players(plyr).health = 100
                    players(plyr).message = STSTR_DQDON
                Else
                    players(plyr).message = STSTR_DQDOFF
                End If

            ' 'fa' cheat for killer fucking arsenal
            ElseIf cht_CheckCheat(cheat_ammonokey, ev.data1) Then
                players(plyr).armorpoints = 200
                players(plyr).armortype = 2

                For i = 0 To NUMWEAPONS - 1
                    players(plyr).weaponowned(i) = True
                Next

                For i = 0 To NUMAMMO - 1
                    players(plyr).ammo(i) = players(plyr).maxammo(i)
                Next

                players(plyr).message = STSTR_FAADDED

            ' 'kfa' cheat for key full ammo
            ElseIf cht_CheckCheat(cheat_ammo, ev.data1) Then
                players(plyr).armorpoints = 200
                players(plyr).armortype = 2

                For i = 0 To NUMWEAPONS - 1
                    players(plyr).weaponowned(i) = True
                Next

                For i = 0 To NUMAMMO - 1
                    players(plyr).ammo(i) = players(plyr).maxammo(i)
                Next

                For i = 0 To NUMCARDS - 1
                    players(plyr).cards(i) = True
                Next

                players(plyr).message = STSTR_KFAADDED

            ' 'mus' cheat for changing music
            ElseIf cht_CheckCheat(cheat_mus, ev.data1) Then
                players(plyr).message = STSTR_MUS
                buf = cht_GetParam(cheat_mus) & Chr$(0) & Chr$(0)

                ' Deviation: a selection below the first song is refused too
                ' (the C passed it on and S_ChangeMusic bombed out).
                If gamemode = commercial Then
                    i = (Asc(Mid$(buf, 1, 1)) - 48) * 10 + Asc(Mid$(buf, 2, 1)) - 48
                    musnum = mus_runnin + i - 1

                    If i > 35 Or i < 1 Then
                        players(plyr).message = STSTR_NOMUS
                    Else
                        S_ChangeMusic musnum, True
                    End If
                Else
                    i = (Asc(Mid$(buf, 1, 1)) - 49) * 9 + Asc(Mid$(buf, 2, 1)) - 49
                    musnum = mus_e1m1 + i

                    If i > 31 Or i < 0 Then
                        players(plyr).message = STSTR_NOMUS
                    Else
                        S_ChangeMusic musnum, True
                    End If
                End If

            Else
                ' Simplified, accepting both "noclip" and "idspispopd".
                ' no clipping mode cheat (C: || short-circuits)
                hit = cht_CheckCheat(cheat_noclip, ev.data1)
                If Not hit Then hit = cht_CheckCheat(cheat_commercial_noclip, ev.data1)
                If hit Then
                    players(plyr).cheats = players(plyr).cheats Xor CF_NOCLIP

                    If (players(plyr).cheats And CF_NOCLIP) <> 0 Then
                        players(plyr).message = STSTR_NCON
                    Else
                        players(plyr).message = STSTR_NCOFF
                    End If
                End If
            End If

            ' 'behold?' power-up cheats
            For i = 0 To 5
                If cht_CheckCheat(cheat_powerup(i), ev.data1) Then
                    If players(plyr).powers(i) = 0 Then
                        P_GivePower plyr, i
                    ElseIf i <> pw_strength Then
                        players(plyr).powers(i) = 1
                    Else
                        players(plyr).powers(i) = 0
                    End If

                    players(plyr).message = STSTR_BEHOLDX
                End If
            Next

            ' 'behold' power-up menu
            If cht_CheckCheat(cheat_powerup(6), ev.data1) Then
                players(plyr).message = STSTR_BEHOLD

            ' 'choppers' invulnerability & chainsaw
            ElseIf cht_CheckCheat(cheat_choppers, ev.data1) Then
                players(plyr).weaponowned(wp_chainsaw) = True
                players(plyr).powers(pw_invulnerability) = 1
                players(plyr).message = STSTR_CHOPPERS

            ' 'mypos' for player position
            ElseIf cht_CheckCheat(cheat_mypos, ev.data1) Then
                mo = players(consoleplayer).mo
                players(plyr).message = "ang=0x" & LCase$(Hex$(mobjs(mo).angle)) & _
                                        ";x,y=(0x" & LCase$(Hex$(mobjs(mo).x)) & _
                                        ",0x" & LCase$(Hex$(mobjs(mo).y)) & ")"
            End If
        End If

        ' 'clev' change-level cheat
        If cht_CheckCheat(cheat_clev, ev.data1) Then
            buf = cht_GetParam(cheat_clev) & Chr$(0) & Chr$(0)

            ' Deviation: the C set epsd = 0 for DOOM 2, so the epsd < 1 test
            ' below refused every DOOM 2 level.
            If gamemode = commercial Then
                epsd = 1
                map = (Asc(Mid$(buf, 1, 1)) - 48) * 10 + Asc(Mid$(buf, 2, 1)) - 48
            Else
                epsd = Asc(Mid$(buf, 1, 1)) - 48
                map = Asc(Mid$(buf, 2, 1)) - 48
            End If

            ' Catch invalid maps.
            If epsd < 1 Then Exit Function

            If map < 1 Then Exit Function

            ' Ohmygod - this is not going to work.
            If gamemode = retail And (epsd > 4 Or map > 9) Then Exit Function

            If gamemode = registered And (epsd > 3 Or map > 9) Then Exit Function

            If gamemode = shareware And (epsd > 1 Or map > 9) Then Exit Function

            If gamemode = commercial And (epsd > 1 Or map > 34) Then Exit Function

            ' So be it.
            players(plyr).message = STSTR_CLEV
            G_DeferedInitNew gameskill, epsd, map
        End If
    End If

    ST_Responder = False
End Function

Private Function ST_calcPainOffset() As Long
    Dim health As Long

    health = players(plyr).health
    If health > 100 Then health = 100

    If health <> oldhealth Then
        lastcalc = ST_FACESTRIDE * (((100 - health) * ST_NUMPAINFACES) \ 101)
        oldhealth = health
    End If
    ST_calcPainOffset = lastcalc
End Function

' This is a not-very-pretty routine which handles the face states and
' their timing.  The precedence of expressions is:
'  dead > evil grin > turned head > straight ahead
Private Sub ST_updateFaceWidget()
    Dim i As Long, badguyangle As Long, diffang As Long, turnright As Boolean
    Dim doevilgrin As Boolean, mo As Long, attacker As Long

    mo = players(plyr).mo
    attacker = players(plyr).attacker

    If priority < 10 Then
        ' dead
        If players(plyr).health = 0 Then
            priority = 9
            st_faceindex = ST_DEADFACE
            st_facecount = 1
        End If
    End If

    If priority < 9 Then
        If players(plyr).bonuscount <> 0 Then
            ' picking up bonus
            doevilgrin = False

            For i = 0 To NUMWEAPONS - 1
                If oldweaponsowned(i) <> players(plyr).weaponowned(i) Then
                    doevilgrin = True
                    oldweaponsowned(i) = players(plyr).weaponowned(i)
                End If
            Next
            If doevilgrin Then
                ' evil grin if just picked up weapon
                priority = 8
                st_facecount = ST_EVILGRINCOUNT
                st_faceindex = ST_calcPainOffset() + ST_EVILGRINOFFSET
            End If
        End If
    End If

    If priority < 8 Then
        If players(plyr).damagecount <> 0 And attacker <> 0 And attacker <> mo Then
            ' being attacked
            priority = 7

            If players(plyr).health - st_oldhealth > ST_MUCHPAIN Then
                st_facecount = ST_TURNCOUNT
                st_faceindex = ST_calcPainOffset() + ST_OUCHOFFSET
            Else
                badguyangle = R_PointToAngle2(mobjs(mo).x, mobjs(mo).y, mobjs(attacker).x, mobjs(attacker).y)

                If UGT(badguyangle, mobjs(mo).angle) Then
                    ' whether right or left
                    diffang = badguyangle - mobjs(mo).angle
                    turnright = UGT(diffang, ANG180)
                Else
                    ' whether left or right
                    diffang = mobjs(mo).angle - badguyangle
                    turnright = ULE(diffang, ANG180)
                End If                           ' confusing, aint it?

                st_facecount = ST_TURNCOUNT
                st_faceindex = ST_calcPainOffset()

                If ULT(diffang, ANG45) Then
                    ' head-on
                    st_faceindex = st_faceindex + ST_RAMPAGEOFFSET
                ElseIf turnright Then
                    ' turn face right
                    st_faceindex = st_faceindex + ST_TURNOFFSET
                Else
                    ' turn face left
                    st_faceindex = st_faceindex + ST_TURNOFFSET + 1
                End If
            End If
        End If
    End If

    If priority < 7 Then
        ' getting hurt because of your own damn stupidity
        If players(plyr).damagecount <> 0 Then
            If players(plyr).health - st_oldhealth > ST_MUCHPAIN Then
                priority = 7
                st_facecount = ST_TURNCOUNT
                st_faceindex = ST_calcPainOffset() + ST_OUCHOFFSET
            Else
                priority = 6
                st_facecount = ST_TURNCOUNT
                st_faceindex = ST_calcPainOffset() + ST_RAMPAGEOFFSET
            End If
        End If
    End If

    If priority < 6 Then
        ' rapid firing
        If players(plyr).attackdown Then
            If lastattackdown = -1 Then
                lastattackdown = ST_RAMPAGEDELAY
            Else
                lastattackdown = lastattackdown - 1
                If lastattackdown = 0 Then
                    priority = 5
                    st_faceindex = ST_calcPainOffset() + ST_RAMPAGEOFFSET
                    st_facecount = 1
                    lastattackdown = 1
                End If
            End If
        Else
            lastattackdown = -1
        End If
    End If

    If priority < 5 Then
        ' invulnerability
        If (players(plyr).cheats And CF_GODMODE) <> 0 Or players(plyr).powers(pw_invulnerability) <> 0 Then
            priority = 4

            st_faceindex = ST_GODFACE
            st_facecount = 1
        End If
    End If

    ' look left or look right if the facecount has timed out
    If st_facecount = 0 Then
        st_faceindex = ST_calcPainOffset() + (st_randomnumber Mod 3)
        st_facecount = ST_STRAIGHTFACECOUNT
        priority = 0
    End If

    st_facecount = st_facecount - 1
End Sub

Private Sub ST_updateWidgets()
    Dim i As Long

    ' must redirect the pointer if the ready weapon has changed.
    If weaponinfo(players(plyr).readyweapon).ammo = am_noammo Then
        st_readyammo = -1                        ' largeammo = 1994 means "n/a"
    Else
        st_readyammo = weaponinfo(players(plyr).readyweapon).ammo
    End If
    w_ready.data = players(plyr).readyweapon

    ' update keycard multiple widgets
    For i = 0 To 2
        If players(plyr).cards(i) Then keyboxes(i) = i Else keyboxes(i) = -1

        If players(plyr).cards(i + 3) Then keyboxes(i) = i + 3
    Next

    ' refresh everything if this is him coming back to life
    ST_updateFaceWidget

    ' used by the w_armsbg widget
    st_notdeathmatch = (deathmatch = 0)

    ' used by w_arms[] widgets
    st_armson = st_statusbaron And (deathmatch = 0)

    ' used by w_frags widget
    st_fragson = (deathmatch <> 0) And st_statusbaron
    st_fragscount = 0

    For i = 0 To MAXPLAYERS - 1
        If i <> consoleplayer Then
            st_fragscount = st_fragscount + players(plyr).frags(i)
        Else
            st_fragscount = st_fragscount - players(plyr).frags(i)
        End If
    Next

    ' get rid of chat window if up because of message
    st_msgcounter = st_msgcounter - 1
    If st_msgcounter = 0 Then st_chat = st_oldchat
End Sub

Public Sub ST_Ticker()
    st_clock = st_clock + 1
    st_randomnumber = M_Random()
    ST_updateWidgets
    st_oldhealth = players(plyr).health
End Sub

Private Sub ST_doPaletteStuff()
    Dim palette As Long, cnt As Long, bzc As Long

    cnt = players(plyr).damagecount

    If players(plyr).powers(pw_strength) <> 0 Then
        ' slowly fade the berzerk out
        bzc = 12 - (players(plyr).powers(pw_strength) \ 64)

        If bzc > cnt Then cnt = bzc
    End If

    If cnt <> 0 Then
        palette = (cnt + 7) \ 8

        If palette >= NUMREDPALS Then palette = NUMREDPALS - 1

        palette = palette + STARTREDPALS

    ElseIf players(plyr).bonuscount <> 0 Then
        palette = (players(plyr).bonuscount + 7) \ 8

        If palette >= NUMBONUSPALS Then palette = NUMBONUSPALS - 1

        palette = palette + STARTBONUSPALS

    ElseIf players(plyr).powers(pw_ironfeet) > 4 * 32 Or (players(plyr).powers(pw_ironfeet) And 8) <> 0 Then
        palette = RADIATIONPAL
    Else
        palette = 0
    End If

    If palette <> st_palette Then
        st_palette = palette
        I_SetPalette W_CacheLumpNum(lu_palette) + palette * 768
    End If
End Sub

Private Sub ST_drawWidgets(ByVal refresh As Boolean)
    Dim i As Long, n As Long

    ' used by w_arms[] widgets
    st_armson = st_statusbaron And (deathmatch = 0)

    ' used by w_frags widget
    st_fragson = (deathmatch <> 0) And st_statusbaron

    If st_readyammo < 0 Then n = 1994 Else n = players(plyr).ammo(st_readyammo)
    STlib_updateNum w_ready, n, st_statusbaron, refresh

    For i = 0 To 3
        STlib_updateNum w_ammo(i), players(plyr).ammo(i), st_statusbaron, refresh
        STlib_updateNum w_maxammo(i), players(plyr).maxammo(i), st_statusbaron, refresh
    Next

    STlib_updatePercent w_health, players(plyr).health, st_statusbaron, refresh
    STlib_updatePercent w_armor, players(plyr).armorpoints, st_statusbaron, refresh

    STlib_updateBinIcon w_armsbg, st_notdeathmatch, st_statusbaron, refresh

    For i = 0 To 5
        If players(plyr).weaponowned(i + 1) Then n = 1 Else n = 0
        STlib_updateMultIcon w_arms(i), n, st_armson, refresh
    Next

    STlib_updateMultIcon w_faces, st_faceindex, st_statusbaron, refresh

    For i = 0 To 2
        STlib_updateMultIcon w_keyboxes(i), keyboxes(i), st_statusbaron, refresh
    Next

    STlib_updateNum w_frags, st_fragscount, st_fragson, refresh
End Sub

Private Sub ST_doRefresh()
    st_firsttime = False

    ' draw status bar background to off-screen buff
    ST_refreshBackground

    ' and refresh all widgets
    ST_drawWidgets True
End Sub

Private Sub ST_diffDraw()
    ' update all widgets
    ST_drawWidgets False
End Sub

Public Sub ST_Drawer(ByVal fullscreen As Boolean, ByVal refresh As Boolean)
    st_statusbaron = (Not fullscreen) Or automapactive
    st_firsttime = st_firsttime Or refresh

    ' Do red-/gold-shifts from damage/items
    ST_doPaletteStuff

    ' If just after ST_Start(), refresh all
    If st_firsttime Then
        ST_doRefresh
    Else
        ' Otherwise, update as little as possible
        ST_diffDraw
    End If
End Sub

Private Sub ST_loadGraphics()
    Dim i As Long, j As Long, facenum As Long

    ' Load the numbers, tall and short
    For i = 0 To 9
        tallnum(i) = W_GetNumForName("STTNUM" & i)
        shortnum(i) = W_GetNumForName("STYSNUM" & i)
    Next

    ' Load percent key.
    'Note: why not load STMINUS here, too?
    tallpercent = W_GetNumForName("STTPRCNT")

    ' key cards
    For i = 0 To NUMCARDS - 1
        keys(i) = W_GetNumForName("STKEYS" & i)
    Next

    ' arms background
    armsbg = W_GetNumForName("STARMS")

    ' arms ownership widgets
    For i = 0 To 5
        arms(i * 2) = W_GetNumForName("STGNUM" & (i + 2))   ' gray #
        arms(i * 2 + 1) = shortnum(i + 2)                    ' yellow #
    Next

    ' face backgrounds for different color players
    faceback = W_GetNumForName("STFB" & consoleplayer)

    ' status bar background bits
    sbar = W_GetNumForName("STBAR")

    ' face states
    facenum = 0
    For i = 0 To ST_NUMPAINFACES - 1
        For j = 0 To ST_NUMSTRAIGHTFACES - 1
            faces(facenum) = W_GetNumForName("STFST" & i & j)
            facenum = facenum + 1
        Next
        faces(facenum) = W_GetNumForName("STFTR" & i & "0")      ' turn right
        faces(facenum + 1) = W_GetNumForName("STFTL" & i & "0")  ' turn left
        faces(facenum + 2) = W_GetNumForName("STFOUCH" & i)      ' ouch!
        faces(facenum + 3) = W_GetNumForName("STFEVL" & i)       ' evil grin ;)
        faces(facenum + 4) = W_GetNumForName("STFKILL" & i)      ' pissed off
        facenum = facenum + 5
    Next
    faces(facenum) = W_GetNumForName("STFGOD0")
    faces(facenum + 1) = W_GetNumForName("STFDEAD0")
End Sub

Private Sub ST_loadData()
    lu_palette = W_GetNumForName("PLAYPAL")
    ST_loadGraphics
End Sub

Private Sub ST_initData()
    Dim i As Long

    st_firsttime = True
    plyr = consoleplayer

    st_clock = 0
    st_chatstate = StartChatState
    st_gamestate = FirstPersonState

    st_statusbaron = True
    st_chat = False
    st_oldchat = False
    st_cursoron = False

    st_faceindex = 0
    st_palette = -1

    st_oldhealth = -1

    For i = 0 To NUMWEAPONS - 1
        oldweaponsowned(i) = players(plyr).weaponowned(i)
    Next

    For i = 0 To 2
        keyboxes(i) = -1
    Next

    STlib_init
End Sub

Private Sub ST_createWidgets()
    Dim i As Long

    ' ready weapon ammo
    If weaponinfo(players(plyr).readyweapon).ammo = am_noammo Then
        st_readyammo = -1
    Else
        st_readyammo = weaponinfo(players(plyr).readyweapon).ammo
    End If
    STlib_initNum w_ready, ST_AMMOX, ST_AMMOY, tallnum, ST_AMMOWIDTH

    ' the last weapon type
    w_ready.data = players(plyr).readyweapon

    ' health percentage
    STlib_initPercent w_health, ST_HEALTHX, ST_HEALTHY, tallnum, tallpercent

    ' arms background
    STlib_initBinIcon w_armsbg, ST_ARMSBGX, ST_ARMSBGY, armsbg

    ' weapons owned
    For i = 0 To 5
        STlib_initMultIcon w_arms(i), ST_ARMSX + (i Mod 3) * ST_ARMSXSPACE, ST_ARMSY + (i \ 3) * ST_ARMSYSPACE, _
                           arms, i * 2, 2
    Next

    ' frags sum
    STlib_initNum w_frags, ST_FRAGSX, ST_FRAGSY, tallnum, ST_FRAGSWIDTH

    ' faces
    STlib_initMultIcon w_faces, ST_FACESX, ST_FACESY, faces, 0, ST_NUMFACES

    ' armor percentage - should be colored later
    STlib_initPercent w_armor, ST_ARMORX, ST_ARMORY, tallnum, tallpercent

    ' keyboxes 0-2
    STlib_initMultIcon w_keyboxes(0), ST_KEY0X, ST_KEY0Y, keys, 0, NUMCARDS
    STlib_initMultIcon w_keyboxes(1), ST_KEY1X, ST_KEY1Y, keys, 0, NUMCARDS
    STlib_initMultIcon w_keyboxes(2), ST_KEY2X, ST_KEY2Y, keys, 0, NUMCARDS

    ' ammo count (all four kinds)
    STlib_initNum w_ammo(0), ST_AMMO0X, ST_AMMO0Y, shortnum, ST_AMMO0WIDTH
    STlib_initNum w_ammo(1), ST_AMMO1X, ST_AMMO1Y, shortnum, ST_AMMO0WIDTH
    STlib_initNum w_ammo(2), ST_AMMO2X, ST_AMMO2Y, shortnum, ST_AMMO0WIDTH
    STlib_initNum w_ammo(3), ST_AMMO3X, ST_AMMO3Y, shortnum, ST_AMMO0WIDTH

    ' max ammo count (all four kinds)
    STlib_initNum w_maxammo(0), ST_MAXAMMO0X, ST_MAXAMMO0Y, shortnum, ST_MAXAMMO0WIDTH
    STlib_initNum w_maxammo(1), ST_MAXAMMO1X, ST_MAXAMMO1Y, shortnum, ST_MAXAMMO0WIDTH
    STlib_initNum w_maxammo(2), ST_MAXAMMO2X, ST_MAXAMMO2Y, shortnum, ST_MAXAMMO0WIDTH
    STlib_initNum w_maxammo(3), ST_MAXAMMO3X, ST_MAXAMMO3Y, shortnum, ST_MAXAMMO0WIDTH
End Sub

Public Sub ST_Start()
    If Not st_stopped Then ST_Stop

    ST_initData
    ST_createWidgets
    st_stopped = False
End Sub

Public Sub ST_Stop()
    If st_stopped Then Exit Sub

    I_SetPalette W_CacheLumpNum(lu_palette)

    st_stopped = True
End Sub

' (screens(4), which the C allocated here, is part of screens().)
Public Sub ST_Init()
    veryfirsttime = 0
    st_stopped = True
    oldhealth = -1
    lastattackdown = -1
    st_oldhealth = -1
    InitCheats
    ST_loadData
End Sub
