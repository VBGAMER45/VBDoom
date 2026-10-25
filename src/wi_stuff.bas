Attribute VB_Name = "wi_stuff"
' wi_stuff.bas - intermission/stat screens.  Different between registered
' DOOM (1994) and Ultimate DOOM - Final edition (retail, 1995?)
' (wi_stuff.c).
'
' WI_Start copies the wbstartstruct (the C kept a pointer to wminfo; g_game
' does not read the fields the intermission changes).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' This is supposedly ignored for commercial release (aka DOOM II), which had
' 34 maps in one episode.  So there.
Private Const NUMEPISODES As Long = 4
Private Const NUMMAPS As Long = 9

' GLOBAL LOCATIONS
Private Const WI_TITLEY As Long = 2
Private Const WI_SPACINGY As Long = 33

' SINGPLE-PLAYER STUFF
Private Const SP_STATSX As Long = 50
Private Const SP_STATSY As Long = 50

Private Const SP_TIMEX As Long = 16
Private Const SP_TIMEY As Long = SCREENHEIGHT - 32

' NET GAME STUFF
Private Const NG_STATSY As Long = 50
Private Const NG_SPACINGX As Long = 64

' DEATHMATCH STUFF
Private Const DM_MATRIXX As Long = 42
Private Const DM_MATRIXY As Long = 68
Private Const DM_SPACINGX As Long = 40
Private Const DM_TOTALSX As Long = 269
Private Const DM_KILLERSX As Long = 10
Private Const DM_KILLERSY As Long = 100
Private Const DM_VICTIMSX As Long = 5
Private Const DM_VICTIMSY As Long = 50

' animenum_t
Private Const ANIM_ALWAYS As Long = 0
Private Const ANIM_RANDOM As Long = 1
Private Const ANIM_LEVEL As Long = 2

' stateenum_t
Private Const NoState As Long = -1
Private Const StatCount As Long = 0
Private Const ShowNextLoc As Long = 1

' Animation.  There is another anim_t used in p_spec.
Private Type wianim_t
    typ As Long                                  ' animenum_t
    period As Long                               ' period in tics between animations
    nanims As Long                               ' number of animation frames
    locx As Long                                 ' location of animation (C: point_t loc)
    locy As Long
    data1 As Long                                ' ALWAYS: n/a, RANDOM: period deviation (<256), LEVEL: level
    data2 As Long                                ' ALWAYS: n/a, RANDOM: random base period, LEVEL: n/a
    p(0 To 2) As Long                            ' actual graphics for frames of animations
    ' following must be initialized to zero before use!
    nexttic As Long                              ' next value of bcnt (used in conjunction with period)
    lastdrawn As Long                            ' last drawn animation frame
    ctr As Long                                  ' next frame number to animate
    state As Long                                ' used by RANDOM and LEVEL when animating
End Type

Private Const FB As Long = 0

' States for single-player
Private Const SP_KILLS As Long = 0
Private Const SP_ITEMS As Long = 2
Private Const SP_SECRET As Long = 4
Private Const SP_FRAGS As Long = 6
Private Const SP_TIME As Long = 8
Private Const SP_PAUSE As Long = 1

' in seconds
Private Const SHOWNEXTLOCDELAY As Long = 4

' location of each level on the world maps (C: point_t lnodes[][])
Private lnodex(0 To 2, 0 To 8) As Long
Private lnodey(0 To 2, 0 To 8) As Long

' Animation locations for the episodes (C: epsd0animinfo ... and anims[])
Private anims(0 To 2, 0 To 9) As wianim_t
Private NUMANIMS(0 To 2) As Long
Private tablesInit As Boolean

' used to accelerate or skip a stage
Private acceleratestage As Long

' wbs->pnum
Private myplayer As Long                         ' wbs->pnum (C: me)

' specifies current state
Private state As Long

' contains information passed into intermission
Private wbs As wbstartstruct_t                   ' (plrs = wbs.plyr)

' used for general timing
Private cnt As Long

' used for timing of background animation
Private bcnt As Long

' signals to refresh everything for one frame
Private firstrefresh As Long

Private cnt_kills(0 To 3) As Long
Private cnt_items(0 To 3) As Long
Private cnt_secret(0 To 3) As Long
Private cnt_time As Long
Private cnt_par As Long
Private cnt_pause As Long

' # of commercial levels
Private NUMCMAPS As Long

' GRAPHICS (lump numbers)

' background (map of levels).
Private bg As Long

' You Are Here graphic
Private yah(0 To 1) As Long

' splat (twice: WI_drawOnLnode may look at a second patch)
Private splat(0 To 1) As Long

' %, : graphics
Private percent As Long
Private colon As Long

' 0-9 graphic
Private num(0 To 9) As Long

' minus sign
Private wiminus As Long

' "Finished!" graphics
Private finished As Long

' "Entering" graphic
Private entering As Long

' "secret"
Private sp_secretpic As Long                     ' (C: sp_secret)

' "Kills", "Scrt", "Items", "Frags"
Private kills As Long
Private secret As Long
Private items As Long
Private frags As Long

' Time sucks.
Private timepic As Long                          ' C: time
Private par As Long
Private sucks As Long

' "killers", "victims"
Private killers As Long
Private victims As Long

' "Total", your face, your dead face
Private total As Long
Private star As Long
Private bstar As Long

' "red P[1..MAXPLAYERS]"
Private p(0 To 3) As Long

' "gray P[1..MAXPLAYERS]"
Private bp(0 To 3) As Long

' Name graphics of each level (centered)
Private lnames(0 To 31) As Long

Private snl_pointeron As Boolean

Private dm_state As Long
Private dm_frags(0 To 3, 0 To 3) As Long
Private dm_totals(0 To 3) As Long

Private cnt_frags(0 To 3) As Long
Private dofrags As Long
Private ng_state As Long

Private sp_state As Long

Private Sub SetNode(ByVal e As Long, ByVal m As Long, ByVal x As Long, ByVal y As Long)
    lnodex(e, m) = x
    lnodey(e, m) = y
End Sub

Private Sub SetAnim(ByVal e As Long, ByVal i As Long, ByVal typ As Long, ByVal period As Long, _
                    ByVal nanims As Long, ByVal x As Long, ByVal y As Long, ByVal data1 As Long)
    anims(e, i).typ = typ
    anims(e, i).period = period
    anims(e, i).nanims = nanims
    anims(e, i).locx = x
    anims(e, i).locy = y
    anims(e, i).data1 = data1
End Sub

' the C static tables
Private Sub WI_InitTables()
    Dim i As Long

    If tablesInit Then Exit Sub
    tablesInit = True

    ' Episode 0 World Map
    SetNode 0, 0, 185, 164: SetNode 0, 1, 148, 143: SetNode 0, 2, 69, 122
    SetNode 0, 3, 209, 102: SetNode 0, 4, 116, 89: SetNode 0, 5, 166, 55
    SetNode 0, 6, 71, 56: SetNode 0, 7, 135, 29: SetNode 0, 8, 71, 24
    ' Episode 1 World Map should go here
    SetNode 1, 0, 254, 25: SetNode 1, 1, 97, 50: SetNode 1, 2, 188, 64
    SetNode 1, 3, 128, 78: SetNode 1, 4, 214, 92: SetNode 1, 5, 133, 130
    SetNode 1, 6, 208, 136: SetNode 1, 7, 148, 140: SetNode 1, 8, 235, 158
    ' Episode 2 World Map should go here
    SetNode 2, 0, 156, 168: SetNode 2, 1, 48, 154: SetNode 2, 2, 174, 95
    SetNode 2, 3, 265, 75: SetNode 2, 4, 130, 48: SetNode 2, 5, 279, 23
    SetNode 2, 6, 198, 48: SetNode 2, 7, 140, 25: SetNode 2, 8, 281, 136

    ' Animation locations for episode 0 (1).  Using patches saves a lot of
    ' space, as they replace 320x200 full screen frames.
    SetAnim 0, 0, ANIM_ALWAYS, TICRATE \ 3, 3, 224, 104, 0
    SetAnim 0, 1, ANIM_ALWAYS, TICRATE \ 3, 3, 184, 160, 0
    SetAnim 0, 2, ANIM_ALWAYS, TICRATE \ 3, 3, 112, 136, 0
    SetAnim 0, 3, ANIM_ALWAYS, TICRATE \ 3, 3, 72, 112, 0
    SetAnim 0, 4, ANIM_ALWAYS, TICRATE \ 3, 3, 88, 96, 0
    SetAnim 0, 5, ANIM_ALWAYS, TICRATE \ 3, 3, 64, 48, 0
    SetAnim 0, 6, ANIM_ALWAYS, TICRATE \ 3, 3, 192, 40, 0
    SetAnim 0, 7, ANIM_ALWAYS, TICRATE \ 3, 3, 136, 16, 0
    SetAnim 0, 8, ANIM_ALWAYS, TICRATE \ 3, 3, 80, 16, 0
    SetAnim 0, 9, ANIM_ALWAYS, TICRATE \ 3, 3, 64, 24, 0
    NUMANIMS(0) = 10

    For i = 0 To 6
        SetAnim 1, i, ANIM_LEVEL, TICRATE \ 3, 1, 128, 136, i + 1
    Next
    SetAnim 1, 7, ANIM_LEVEL, TICRATE \ 3, 3, 192, 144, 8
    SetAnim 1, 8, ANIM_LEVEL, TICRATE \ 3, 1, 128, 136, 8
    NUMANIMS(1) = 9

    SetAnim 2, 0, ANIM_ALWAYS, TICRATE \ 3, 3, 104, 168, 0
    SetAnim 2, 1, ANIM_ALWAYS, TICRATE \ 3, 3, 40, 136, 0
    SetAnim 2, 2, ANIM_ALWAYS, TICRATE \ 3, 3, 160, 96, 0
    SetAnim 2, 3, ANIM_ALWAYS, TICRATE \ 3, 3, 104, 80, 0
    SetAnim 2, 4, ANIM_ALWAYS, TICRATE \ 3, 3, 120, 32, 0
    SetAnim 2, 5, ANIM_ALWAYS, TICRATE \ 4, 3, 40, 0, 0
    NUMANIMS(2) = 6
End Sub

' CODE

' slam background
Private Sub WI_slamBackground()
    CopyMemory screens(0), screens(SCREENSIZE), SCREENWIDTH * SCREENHEIGHT
    V_MarkRect 0, 0, SCREENWIDTH, SCREENHEIGHT
End Sub

' The ticker is used to detect keys because of timing issues in netgames.
Public Function WI_Responder(ev As event_t) As Boolean
    WI_Responder = False
End Function

' Draws "<Levelname> Finished!"
Private Sub WI_drawLF()
    Dim y As Long

    y = WI_TITLEY

    ' draw <LevelName>
    V_DrawPatch (SCREENWIDTH - PatchWidth(lnames(wbs.last))) \ 2, y, FB, lnames(wbs.last)

    ' draw "Finished!"
    y = y + (5 * PatchHeight(lnames(wbs.last))) \ 4

    V_DrawPatch (SCREENWIDTH - PatchWidth(finished)) \ 2, y, FB, finished
End Sub

' Draws "Entering <LevelName>"
Private Sub WI_drawEL()
    Dim y As Long

    y = WI_TITLEY

    ' draw "Entering"
    V_DrawPatch (SCREENWIDTH - PatchWidth(entering)) \ 2, y, FB, entering

    ' draw level
    y = y + (5 * PatchHeight(lnames(wbs.nxt))) \ 4

    V_DrawPatch (SCREENWIDTH - PatchWidth(lnames(wbs.nxt))) \ 2, y, FB, lnames(wbs.nxt)
End Sub

Private Sub WI_drawOnLnode(ByVal n As Long, c() As Long)
    Dim i As Long, left_ As Long, top As Long, right_ As Long, bottom As Long, fits As Boolean

    fits = False
    i = 0
    Do
        left_ = lnodex(wbs.epsd, n) - PatchLeftOffset(c(i))
        top = lnodey(wbs.epsd, n) - PatchTopOffset(c(i))
        right_ = left_ + PatchWidth(c(i))
        bottom = top + PatchHeight(c(i))

        If left_ >= 0 And right_ < SCREENWIDTH And top >= 0 And bottom < SCREENHEIGHT Then
            fits = True
        Else
            i = i + 1
        End If
    Loop While Not fits And i <> 2

    If fits And i < 2 Then
        V_DrawPatch lnodex(wbs.epsd, n), lnodey(wbs.epsd, n), FB, c(i)
    Else
        ' DEBUG
        I_Printf "Could not place patch on level " & (n + 1)
    End If
End Sub

Private Sub WI_initAnimatedBack()
    Dim i As Long, e As Long

    If gamemode = commercial Then Exit Sub

    If wbs.epsd > 2 Then Exit Sub

    e = wbs.epsd
    For i = 0 To NUMANIMS(e) - 1
        ' init variables
        anims(e, i).ctr = -1

        ' specify the next time to draw it
        If anims(e, i).typ = ANIM_ALWAYS Then
            anims(e, i).nexttic = bcnt + 1 + (M_Random() Mod anims(e, i).period)
        ElseIf anims(e, i).typ = ANIM_RANDOM Then
            anims(e, i).nexttic = bcnt + 1 + anims(e, i).data2 + (M_Random() Mod anims(e, i).data1)
        ElseIf anims(e, i).typ = ANIM_LEVEL Then
            anims(e, i).nexttic = bcnt + 1
        End If
    Next
End Sub

Private Sub WI_updateAnimatedBack()
    Dim i As Long, e As Long

    If gamemode = commercial Then Exit Sub

    If wbs.epsd > 2 Then Exit Sub

    e = wbs.epsd
    For i = 0 To NUMANIMS(e) - 1
        If bcnt = anims(e, i).nexttic Then
            Select Case anims(e, i).typ
                Case ANIM_ALWAYS
                    anims(e, i).ctr = anims(e, i).ctr + 1
                    If anims(e, i).ctr >= anims(e, i).nanims Then anims(e, i).ctr = 0
                    anims(e, i).nexttic = bcnt + anims(e, i).period

                Case ANIM_RANDOM
                    anims(e, i).ctr = anims(e, i).ctr + 1
                    If anims(e, i).ctr = anims(e, i).nanims Then
                        anims(e, i).ctr = -1
                        anims(e, i).nexttic = bcnt + anims(e, i).data2 + (M_Random() Mod anims(e, i).data1)
                    Else
                        anims(e, i).nexttic = bcnt + anims(e, i).period
                    End If

                Case ANIM_LEVEL
                    ' gawd-awful hack for level anims
                    If Not (state = StatCount And i = 7) And wbs.nxt = anims(e, i).data1 Then
                        anims(e, i).ctr = anims(e, i).ctr + 1
                        If anims(e, i).ctr = anims(e, i).nanims Then anims(e, i).ctr = anims(e, i).ctr - 1
                        anims(e, i).nexttic = bcnt + anims(e, i).period
                    End If
            End Select
        End If
    Next
End Sub

' Deviation: the C tested "if (commercial)", the enum constant, so the
' animations were never drawn; DOOM 2 is meant here.
Private Sub WI_drawAnimatedBack()
    Dim i As Long, e As Long

    If gamemode = commercial Then Exit Sub

    If wbs.epsd > 2 Then Exit Sub

    e = wbs.epsd
    For i = 0 To NUMANIMS(e) - 1
        If anims(e, i).ctr >= 0 Then
            V_DrawPatch anims(e, i).locx, anims(e, i).locy, FB, anims(e, i).p(anims(e, i).ctr)
        End If
    Next
End Sub

' Draws a number.
' If digits > 0, then use that many digits minimum,
'  otherwise only use as many as necessary.
' Returns new x position.
Private Function WI_drawNum(ByVal x As Long, ByVal y As Long, ByVal n As Long, ByVal digits As Long) As Long
    Dim fontwidth As Long, neg As Boolean, temp As Long

    fontwidth = PatchWidth(num(0))

    If digits < 0 Then
        If n = 0 Then
            ' make variable-length zeros 1 digit long
            digits = 1
        Else
            ' figure out # of digits in #
            digits = 0
            temp = n
            Do While temp <> 0
                temp = temp \ 10
                digits = digits + 1
            Loop
        End If
    End If

    neg = n < 0
    If neg Then n = -n

    ' if non-number, do not draw it
    If n = 1994 Then
        WI_drawNum = 0
        Exit Function
    End If

    ' draw the new number
    Do While digits <> 0
        digits = digits - 1
        x = x - fontwidth
        V_DrawPatch x, y, FB, num(n Mod 10)
        n = n \ 10
    Loop

    ' draw a minus sign if necessary
    If neg Then
        x = x - 8
        V_DrawPatch x, y, FB, wiminus
    End If

    WI_drawNum = x
End Function

Private Sub WI_drawPercent(ByVal x As Long, ByVal y As Long, ByVal pc As Long)
    If pc < 0 Then Exit Sub

    V_DrawPatch x, y, FB, percent
    WI_drawNum x, y, pc, -1
End Sub

' Display level completion time and par, or "sucks" message if overflow.
Private Sub WI_drawTime(ByVal x As Long, ByVal y As Long, ByVal t As Long)
    Dim dv As Long, n As Long

    If t < 0 Then Exit Sub

    If t <= 61 * 59 Then
        dv = 1
        Do
            n = (t \ dv) Mod 60
            x = WI_drawNum(x, y, n, 2) - PatchWidth(colon)
            dv = dv * 60

            ' draw
            If dv = 60 Or t \ dv <> 0 Then V_DrawPatch x, y, FB, colon
        Loop While t \ dv <> 0
    Else
        ' "sucks"
        V_DrawPatch x - PatchWidth(sucks), y, FB, sucks
    End If
End Sub

Private Sub WI_End()
    WI_unloadData
End Sub

Private Sub WI_initNoState()
    state = NoState
    acceleratestage = 0
    cnt = 10
End Sub

Private Sub WI_updateNoState()
    WI_updateAnimatedBack

    cnt = cnt - 1
    If cnt = 0 Then
        WI_End
        G_WorldDone
    End If
End Sub

Private Sub WI_initShowNextLoc()
    state = ShowNextLoc
    acceleratestage = 0
    cnt = SHOWNEXTLOCDELAY * TICRATE

    WI_initAnimatedBack
End Sub

Private Sub WI_updateShowNextLoc()
    WI_updateAnimatedBack

    cnt = cnt - 1
    If cnt = 0 Or acceleratestage <> 0 Then
        WI_initNoState
    Else
        snl_pointeron = (cnt And 31) < 20
    End If
End Sub

Private Sub WI_drawShowNextLoc()
    Dim i As Long, last As Long

    WI_slamBackground

    ' draw animated background
    WI_drawAnimatedBack

    If gamemode <> commercial Then
        If wbs.epsd > 2 Then
            WI_drawEL
            Exit Sub
        End If

        If wbs.last = 8 Then last = wbs.nxt - 1 Else last = wbs.last

        ' draw a splat on taken cities.
        For i = 0 To last
            WI_drawOnLnode i, splat
        Next

        ' splat the secret level?
        If wbs.didsecret Then WI_drawOnLnode 8, splat

        ' draw flashing ptr
        If snl_pointeron Then WI_drawOnLnode wbs.nxt, yah
    End If

    ' draws which level you are entering..
    If gamemode <> commercial Or wbs.nxt <> 30 Then WI_drawEL
End Sub

Private Sub WI_drawNoState()
    snl_pointeron = True
    WI_drawShowNextLoc
End Sub

Private Function WI_fragSum(ByVal playernum As Long) As Long
    Dim i As Long, fr As Long

    fr = 0
    For i = 0 To MAXPLAYERS - 1
        If playeringame(i) And i <> playernum Then fr = fr + wbs.plyr(playernum).frags(i)
    Next

    ' JDC hack - negative frags.
    fr = fr - wbs.plyr(playernum).frags(playernum)
    ' UNUSED if (frags < 0)
    '   frags = 0;

    WI_fragSum = fr
End Function

Private Sub WI_initDeathmatchStats()
    Dim i As Long, j As Long

    state = StatCount
    acceleratestage = 0
    dm_state = 1

    cnt_pause = TICRATE

    For i = 0 To MAXPLAYERS - 1
        If playeringame(i) Then
            For j = 0 To MAXPLAYERS - 1
                If playeringame(j) Then dm_frags(i, j) = 0
            Next
            dm_totals(i) = 0
        End If
    Next

    WI_initAnimatedBack
End Sub

Private Sub WI_updateDeathmatchStats()
    Dim i As Long, j As Long, stillticking As Boolean

    WI_updateAnimatedBack

    If acceleratestage <> 0 And dm_state <> 4 Then
        acceleratestage = 0

        For i = 0 To MAXPLAYERS - 1
            If playeringame(i) Then
                For j = 0 To MAXPLAYERS - 1
                    If playeringame(j) Then dm_frags(i, j) = wbs.plyr(i).frags(j)
                Next
                dm_totals(i) = WI_fragSum(i)
            End If
        Next

        S_StartSound 0, sfx_barexp
        dm_state = 4
    End If

    If dm_state = 2 Then
        If (bcnt And 3) = 0 Then S_StartSound 0, sfx_pistol

        stillticking = False

        For i = 0 To MAXPLAYERS - 1
            If playeringame(i) Then
                For j = 0 To MAXPLAYERS - 1
                    If playeringame(j) Then
                        If dm_frags(i, j) <> wbs.plyr(i).frags(j) Then
                            If wbs.plyr(i).frags(j) < 0 Then
                                dm_frags(i, j) = dm_frags(i, j) - 1
                            Else
                                dm_frags(i, j) = dm_frags(i, j) + 1
                            End If

                            If dm_frags(i, j) > 99 Then dm_frags(i, j) = 99

                            If dm_frags(i, j) < -99 Then dm_frags(i, j) = -99

                            stillticking = True
                        End If
                    End If
                Next
                dm_totals(i) = WI_fragSum(i)

                If dm_totals(i) > 99 Then dm_totals(i) = 99

                If dm_totals(i) < -99 Then dm_totals(i) = -99
            End If
        Next
        If Not stillticking Then
            S_StartSound 0, sfx_barexp
            dm_state = dm_state + 1
        End If

    ElseIf dm_state = 4 Then
        If acceleratestage <> 0 Then
            S_StartSound 0, sfx_slop

            If gamemode = commercial Then
                WI_initNoState
            Else
                WI_initShowNextLoc
            End If
        End If

    ElseIf (dm_state And 1) <> 0 Then
        cnt_pause = cnt_pause - 1
        If cnt_pause = 0 Then
            dm_state = dm_state + 1
            cnt_pause = TICRATE
        End If
    End If
End Sub

Private Sub WI_drawDeathmatchStats()
    Dim i As Long, j As Long, x As Long, y As Long, w As Long, lh As Long

    lh = WI_SPACINGY                             ' line height

    WI_slamBackground

    ' draw animated background
    WI_drawAnimatedBack
    WI_drawLF

    ' draw stat titles (top line)
    V_DrawPatch DM_TOTALSX - PatchWidth(total) \ 2, DM_MATRIXY - WI_SPACINGY + 10, FB, total

    V_DrawPatch DM_KILLERSX, DM_KILLERSY, FB, killers
    V_DrawPatch DM_VICTIMSX, DM_VICTIMSY, FB, victims

    ' draw P?
    x = DM_MATRIXX + DM_SPACINGX
    y = DM_MATRIXY

    For i = 0 To MAXPLAYERS - 1
        If playeringame(i) Then
            V_DrawPatch x - PatchWidth(p(i)) \ 2, DM_MATRIXY - WI_SPACINGY, FB, p(i)

            V_DrawPatch DM_MATRIXX - PatchWidth(p(i)) \ 2, y, FB, p(i)

            If i = myplayer Then
                V_DrawPatch x - PatchWidth(p(i)) \ 2, DM_MATRIXY - WI_SPACINGY, FB, bstar

                V_DrawPatch DM_MATRIXX - PatchWidth(p(i)) \ 2, y, FB, star
            End If
        End If
        x = x + DM_SPACINGX
        y = y + WI_SPACINGY
    Next

    ' draw stats
    y = DM_MATRIXY + 10
    w = PatchWidth(num(0))

    For i = 0 To MAXPLAYERS - 1
        x = DM_MATRIXX + DM_SPACINGX

        If playeringame(i) Then
            For j = 0 To MAXPLAYERS - 1
                If playeringame(j) Then WI_drawNum x + w, y, dm_frags(i, j), 2
                x = x + DM_SPACINGX
            Next
            WI_drawNum DM_TOTALSX + w, y, dm_totals(i), 2
        End If
        y = y + WI_SPACINGY
    Next
End Sub

Private Sub WI_initNetgameStats()
    Dim i As Long

    state = StatCount
    acceleratestage = 0
    ng_state = 1

    cnt_pause = TICRATE

    For i = 0 To MAXPLAYERS - 1
        If playeringame(i) Then
            cnt_kills(i) = 0
            cnt_items(i) = 0
            cnt_secret(i) = 0
            cnt_frags(i) = 0

            dofrags = dofrags + WI_fragSum(i)
        End If
    Next

    If dofrags <> 0 Then dofrags = 1             ' !!dofrags

    WI_initAnimatedBack
End Sub

Private Sub WI_updateNetgameStats()
    Dim i As Long, fsum As Long, stillticking As Boolean

    WI_updateAnimatedBack

    If acceleratestage <> 0 And ng_state <> 10 Then
        acceleratestage = 0

        For i = 0 To MAXPLAYERS - 1
            If playeringame(i) Then
                cnt_kills(i) = (wbs.plyr(i).skills * 100) \ wbs.maxkills
                cnt_items(i) = (wbs.plyr(i).sitems * 100) \ wbs.maxitems
                cnt_secret(i) = (wbs.plyr(i).ssecret * 100) \ wbs.maxsecret

                If dofrags <> 0 Then cnt_frags(i) = WI_fragSum(i)
            End If
        Next
        S_StartSound 0, sfx_barexp
        ng_state = 10
    End If

    If ng_state = 2 Then
        If (bcnt And 3) = 0 Then S_StartSound 0, sfx_pistol

        stillticking = False

        For i = 0 To MAXPLAYERS - 1
            If playeringame(i) Then
                cnt_kills(i) = cnt_kills(i) + 2

                If cnt_kills(i) >= (wbs.plyr(i).skills * 100) \ wbs.maxkills Then
                    cnt_kills(i) = (wbs.plyr(i).skills * 100) \ wbs.maxkills
                Else
                    stillticking = True
                End If
            End If
        Next

        If Not stillticking Then
            S_StartSound 0, sfx_barexp
            ng_state = ng_state + 1
        End If

    ElseIf ng_state = 4 Then
        If (bcnt And 3) = 0 Then S_StartSound 0, sfx_pistol

        stillticking = False

        For i = 0 To MAXPLAYERS - 1
            If playeringame(i) Then
                cnt_items(i) = cnt_items(i) + 2
                If cnt_items(i) >= (wbs.plyr(i).sitems * 100) \ wbs.maxitems Then
                    cnt_items(i) = (wbs.plyr(i).sitems * 100) \ wbs.maxitems
                Else
                    stillticking = True
                End If
            End If
        Next
        If Not stillticking Then
            S_StartSound 0, sfx_barexp
            ng_state = ng_state + 1
        End If

    ElseIf ng_state = 6 Then
        If (bcnt And 3) = 0 Then S_StartSound 0, sfx_pistol

        stillticking = False

        For i = 0 To MAXPLAYERS - 1
            If playeringame(i) Then
                cnt_secret(i) = cnt_secret(i) + 2

                If cnt_secret(i) >= (wbs.plyr(i).ssecret * 100) \ wbs.maxsecret Then
                    cnt_secret(i) = (wbs.plyr(i).ssecret * 100) \ wbs.maxsecret
                Else
                    stillticking = True
                End If
            End If
        Next

        If Not stillticking Then
            S_StartSound 0, sfx_barexp
            If dofrags = 0 Then ng_state = ng_state + 3 Else ng_state = ng_state + 1 ' 1 + 2*!dofrags
        End If

    ElseIf ng_state = 8 Then
        If (bcnt And 3) = 0 Then S_StartSound 0, sfx_pistol

        stillticking = False

        For i = 0 To MAXPLAYERS - 1
            If playeringame(i) Then
                cnt_frags(i) = cnt_frags(i) + 1

                fsum = WI_fragSum(i)
                If cnt_frags(i) >= fsum Then
                    cnt_frags(i) = fsum
                Else
                    stillticking = True
                End If
            End If
        Next

        If Not stillticking Then
            S_StartSound 0, sfx_pldeth
            ng_state = ng_state + 1
        End If

    ElseIf ng_state = 10 Then
        If acceleratestage <> 0 Then
            S_StartSound 0, sfx_sgcock
            If gamemode = commercial Then
                WI_initNoState
            Else
                WI_initShowNextLoc
            End If
        End If

    ElseIf (ng_state And 1) <> 0 Then
        cnt_pause = cnt_pause - 1
        If cnt_pause = 0 Then
            ng_state = ng_state + 1
            cnt_pause = TICRATE
        End If
    End If
End Sub

Private Sub WI_drawNetgameStats()
    Dim i As Long, x As Long, y As Long, pwidth As Long, ng_statsx As Long

    pwidth = PatchWidth(percent)
    ' NG_STATSX
    ng_statsx = 32 + PatchWidth(star) \ 2
    If dofrags = 0 Then ng_statsx = ng_statsx + 32

    WI_slamBackground

    ' draw animated background
    WI_drawAnimatedBack

    WI_drawLF

    ' draw stat titles (top line)
    V_DrawPatch ng_statsx + NG_SPACINGX - PatchWidth(kills), NG_STATSY, FB, kills

    V_DrawPatch ng_statsx + 2 * NG_SPACINGX - PatchWidth(items), NG_STATSY, FB, items

    V_DrawPatch ng_statsx + 3 * NG_SPACINGX - PatchWidth(secret), NG_STATSY, FB, secret

    If dofrags <> 0 Then V_DrawPatch ng_statsx + 4 * NG_SPACINGX - PatchWidth(frags), NG_STATSY, FB, frags

    ' draw stats
    y = NG_STATSY + PatchHeight(kills)

    For i = 0 To MAXPLAYERS - 1
        If playeringame(i) Then
            x = ng_statsx
            V_DrawPatch x - PatchWidth(p(i)), y, FB, p(i)

            If i = myplayer Then V_DrawPatch x - PatchWidth(p(i)), y, FB, star

            x = x + NG_SPACINGX
            WI_drawPercent x - pwidth, y + 10, cnt_kills(i)
            x = x + NG_SPACINGX
            WI_drawPercent x - pwidth, y + 10, cnt_items(i)
            x = x + NG_SPACINGX
            WI_drawPercent x - pwidth, y + 10, cnt_secret(i)
            x = x + NG_SPACINGX

            If dofrags <> 0 Then WI_drawNum x, y + 10, cnt_frags(i), -1

            y = y + WI_SPACINGY
        End If
    Next
End Sub

Private Sub WI_initStats()
    state = StatCount
    acceleratestage = 0
    sp_state = 1
    cnt_kills(0) = -1
    cnt_items(0) = -1
    cnt_secret(0) = -1
    cnt_time = -1
    cnt_par = -1
    cnt_pause = TICRATE

    WI_initAnimatedBack
End Sub

Private Sub WI_updateStats()
    WI_updateAnimatedBack

    If acceleratestage <> 0 And sp_state <> 10 Then
        acceleratestage = 0
        cnt_kills(0) = (wbs.plyr(myplayer).skills * 100) \ wbs.maxkills
        cnt_items(0) = (wbs.plyr(myplayer).sitems * 100) \ wbs.maxitems
        cnt_secret(0) = (wbs.plyr(myplayer).ssecret * 100) \ wbs.maxsecret
        cnt_time = wbs.plyr(myplayer).stime \ TICRATE
        cnt_par = wbs.partime \ TICRATE
        S_StartSound 0, sfx_barexp
        sp_state = 10
    End If

    If sp_state = 2 Then
        cnt_kills(0) = cnt_kills(0) + 2

        If (bcnt And 3) = 0 Then S_StartSound 0, sfx_pistol

        If cnt_kills(0) >= (wbs.plyr(myplayer).skills * 100) \ wbs.maxkills Then
            cnt_kills(0) = (wbs.plyr(myplayer).skills * 100) \ wbs.maxkills
            S_StartSound 0, sfx_barexp
            sp_state = sp_state + 1
        End If

    ElseIf sp_state = 4 Then
        cnt_items(0) = cnt_items(0) + 2

        If (bcnt And 3) = 0 Then S_StartSound 0, sfx_pistol

        If cnt_items(0) >= (wbs.plyr(myplayer).sitems * 100) \ wbs.maxitems Then
            cnt_items(0) = (wbs.plyr(myplayer).sitems * 100) \ wbs.maxitems
            S_StartSound 0, sfx_barexp
            sp_state = sp_state + 1
        End If

    ElseIf sp_state = 6 Then
        cnt_secret(0) = cnt_secret(0) + 2

        If (bcnt And 3) = 0 Then S_StartSound 0, sfx_pistol

        If cnt_secret(0) >= (wbs.plyr(myplayer).ssecret * 100) \ wbs.maxsecret Then
            cnt_secret(0) = (wbs.plyr(myplayer).ssecret * 100) \ wbs.maxsecret
            S_StartSound 0, sfx_barexp
            sp_state = sp_state + 1
        End If

    ElseIf sp_state = 8 Then
        If (bcnt And 3) = 0 Then S_StartSound 0, sfx_pistol

        cnt_time = cnt_time + 3

        If cnt_time >= wbs.plyr(myplayer).stime \ TICRATE Then cnt_time = wbs.plyr(myplayer).stime \ TICRATE

        cnt_par = cnt_par + 3

        If cnt_par >= wbs.partime \ TICRATE Then
            cnt_par = wbs.partime \ TICRATE

            If cnt_time >= wbs.plyr(myplayer).stime \ TICRATE Then
                S_StartSound 0, sfx_barexp
                sp_state = sp_state + 1
            End If
        End If

    ElseIf sp_state = 10 Then
        If acceleratestage <> 0 Then
            S_StartSound 0, sfx_sgcock

            If gamemode = commercial Then
                WI_initNoState
            Else
                WI_initShowNextLoc
            End If
        End If

    ElseIf (sp_state And 1) <> 0 Then
        cnt_pause = cnt_pause - 1
        If cnt_pause = 0 Then
            sp_state = sp_state + 1
            cnt_pause = TICRATE
        End If
    End If
End Sub

Private Sub WI_drawStats()
    Dim lh As Long                               ' line height

    lh = (3 * PatchHeight(num(0))) \ 2

    WI_slamBackground

    ' draw animated background
    WI_drawAnimatedBack

    WI_drawLF

    V_DrawPatch SP_STATSX, SP_STATSY, FB, kills
    WI_drawPercent SCREENWIDTH - SP_STATSX, SP_STATSY, cnt_kills(0)

    V_DrawPatch SP_STATSX, SP_STATSY + lh, FB, items
    WI_drawPercent SCREENWIDTH - SP_STATSX, SP_STATSY + lh, cnt_items(0)

    V_DrawPatch SP_STATSX, SP_STATSY + 2 * lh, FB, sp_secretpic
    WI_drawPercent SCREENWIDTH - SP_STATSX, SP_STATSY + 2 * lh, cnt_secret(0)

    V_DrawPatch SP_TIMEX, SP_TIMEY, FB, timepic
    WI_drawTime SCREENWIDTH \ 2 - SP_TIMEX, SP_TIMEY, cnt_time

    If wbs.epsd < 3 Then
        V_DrawPatch SCREENWIDTH \ 2 + SP_TIMEX, SP_TIMEY, FB, par
        WI_drawTime SCREENWIDTH - SP_TIMEX, SP_TIMEY, cnt_par
    End If
End Sub

Private Sub WI_checkForAccelerate()
    Dim i As Long

    ' check for button presses to skip delays
    For i = 0 To MAXPLAYERS - 1
        If playeringame(i) Then
            If (players(i).cmd.buttons And BT_ATTACK) <> 0 Then
                If Not players(i).attackdown Then acceleratestage = 1
                players(i).attackdown = True
            Else
                players(i).attackdown = False
            End If
            If (players(i).cmd.buttons And BT_USE) <> 0 Then
                If Not players(i).usedown Then acceleratestage = 1
                players(i).usedown = True
            Else
                players(i).usedown = False
            End If
        End If
    Next
End Sub

' Updates stuff each tick
Public Sub WI_Ticker()
    ' counter for general background animation
    bcnt = bcnt + 1

    If bcnt = 1 Then
        ' intermission music
        If gamemode = commercial Then
            S_ChangeMusic mus_dm2int, True
        Else
            S_ChangeMusic mus_inter, True
        End If
    End If

    WI_checkForAccelerate

    Select Case state
        Case StatCount
            If deathmatch <> 0 Then
                WI_updateDeathmatchStats
            ElseIf netgame Then
                WI_updateNetgameStats
            Else
                WI_updateStats
            End If

        Case ShowNextLoc
            WI_updateShowNextLoc

        Case NoState
            WI_updateNoState
    End Select
End Sub

' Deviation: the C chose WIOBJ for netgames by "if (french)", the enum
' constant (always true); most IWADs (Freedoom too) have no WIOBJ.
Private Sub WI_loadData()
    Dim i As Long, j As Long, nm As String, e As Long

    If gamemode = commercial Then
        nm = "INTERPIC"
    Else
        nm = "WIMAP" & wbs.epsd
    End If

    If gamemode = retail Then
        If wbs.epsd = 3 Then nm = "INTERPIC"
    End If

    ' background
    bg = W_GetNumForName(nm)
    V_DrawPatch 0, 0, 1, bg

    If gamemode = commercial Then
        NUMCMAPS = 32
        For i = 0 To NUMCMAPS - 1
            lnames(i) = W_GetNumForName("CWILV" & Format$(i, "00"))
        Next
    Else
        For i = 0 To NUMMAPS - 1
            lnames(i) = W_GetNumForName("WILV" & wbs.epsd & i)
        Next

        ' you are here
        yah(0) = W_GetNumForName("WIURH0")

        ' you are here (alt.)
        yah(1) = W_GetNumForName("WIURH1")

        ' splat
        splat(0) = W_GetNumForName("WISPLAT")
        splat(1) = splat(0)

        If wbs.epsd < 3 Then
            e = wbs.epsd
            For j = 0 To NUMANIMS(e) - 1
                For i = 0 To anims(e, j).nanims - 1
                    ' MONDO HACK!
                    If e <> 1 Or j <> 8 Then
                        ' animations
                        anims(e, j).p(i) = W_GetNumForName("WIA" & e & Format$(j, "00") & Format$(i, "00"))
                    Else
                        ' HACK ALERT!
                        anims(e, j).p(i) = anims(1, 4).p(i)
                    End If
                Next
            Next
        End If
    End If

    ' More hacks on minus sign.
    wiminus = W_GetNumForName("WIMINUS")

    For i = 0 To 9
        ' numbers 0-9
        num(i) = W_GetNumForName("WINUM" & i)
    Next

    ' percent sign
    percent = W_GetNumForName("WIPCNT")

    ' "finished"
    finished = W_GetNumForName("WIF")

    ' "entering"
    entering = W_GetNumForName("WIENTER")

    ' "kills"
    kills = W_GetNumForName("WIOSTK")

    ' "scrt"
    secret = W_GetNumForName("WIOSTS")

    ' "secret"
    sp_secretpic = W_GetNumForName("WISCRT2")

    ' Yuck.
    If language = french Then
        ' "items"
        If netgame And deathmatch = 0 Then
            items = W_GetNumForName("WIOBJ")
        Else
            items = W_GetNumForName("WIOSTI")
        End If
    Else
        items = W_GetNumForName("WIOSTI")
    End If

    ' "frgs"
    frags = W_GetNumForName("WIFRGS")

    ' ":"
    colon = W_GetNumForName("WICOLON")

    ' "time"
    timepic = W_GetNumForName("WITIME")

    ' "sucks"
    sucks = W_GetNumForName("WISUCKS")

    ' "par"
    par = W_GetNumForName("WIPAR")

    ' "killers" (vertical)
    killers = W_GetNumForName("WIKILRS")

    ' "victims" (horiz)
    victims = W_GetNumForName("WIVCTMS")

    ' "total"
    total = W_GetNumForName("WIMSTT")

    ' your face
    star = W_GetNumForName("STFST01")

    ' dead face
    bstar = W_GetNumForName("STFDEAD0")

    For i = 0 To MAXPLAYERS - 1
        ' "1,2,3,4"
        p(i) = W_GetNumForName("STPB" & i)

        ' "1,2,3,4"
        bp(i) = W_GetNumForName("WIBP" & (i + 1))
    Next
End Sub

' (nothing to free: the patches are lump numbers)
Private Sub WI_unloadData()
End Sub

Public Sub WI_Drawer()
    Select Case state
        Case StatCount
            If deathmatch <> 0 Then
                WI_drawDeathmatchStats
            ElseIf netgame Then
                WI_drawNetgameStats
            Else
                WI_drawStats
            End If

        Case ShowNextLoc
            WI_drawShowNextLoc

        Case NoState
            WI_drawNoState
    End Select
End Sub

Private Sub WI_initVariables(wbstartstruct As wbstartstruct_t)
    wbs = wbstartstruct

    acceleratestage = 0
    cnt = 0
    bcnt = 0
    firstrefresh = 1
    myplayer = wbs.pnum

    If wbs.maxkills = 0 Then wbs.maxkills = 1

    If wbs.maxitems = 0 Then wbs.maxitems = 1

    If wbs.maxsecret = 0 Then wbs.maxsecret = 1

    If gamemode <> retail Then
        If wbs.epsd > 2 Then wbs.epsd = wbs.epsd - 3
    End If
End Sub

Public Sub WI_Start(wbstartstruct As wbstartstruct_t)
    WI_InitTables
    WI_initVariables wbstartstruct
    WI_loadData

    If deathmatch <> 0 Then
        WI_initDeathmatchStats
    ElseIf netgame Then
        WI_initNetgameStats
    Else
        WI_initStats
    End If
End Sub
