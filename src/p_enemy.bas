Attribute VB_Name = "p_enemy"
' p_enemy.bas - enemy thinking, AI.  Action Pointer Functions that are
' associated with states/frames (p_enemy.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' dirtype_t
Private Const DI_EAST As Long = 0
Private Const DI_NORTHEAST As Long = 1
Private Const DI_NORTH As Long = 2
Private Const DI_NORTHWEST As Long = 3
Private Const DI_WEST As Long = 4
Private Const DI_SOUTHWEST As Long = 5
Private Const DI_SOUTH As Long = 6
Private Const DI_SOUTHEAST As Long = 7
Private Const DI_NODIR As Long = 8

Private Const TRACEANGLE As Long = &HC000000
Private Const FATSPREAD As Long = ANG90 \ 8
Private Const SKULLSPEED As Long = 20 * FRACUNIT

Private opposite(0 To 8) As Long
Private diags(0 To 3) As Long
Private xspeed(0 To 7) As Long
Private yspeed(0 To 7) As Long
Private tablesInit As Boolean

Private soundtarget As Long

' Arch-vile resurrection
Private corpsehit As Long
Private vileobj As Long
Private viletryx As Long
Private viletryy As Long

' Boss brain
Private braintargets() As Long
Private numbraintargets As Long
Private braintargeton As Long
Private brainEasy As Long

Private Sub InitTables()
    Dim v As Variant, i As Long
    v = Array(DI_WEST, DI_SOUTHWEST, DI_SOUTH, DI_SOUTHEAST, DI_EAST, DI_NORTHEAST, DI_NORTH, DI_NORTHWEST, DI_NODIR)
    For i = 0 To 8: opposite(i) = v(i): Next
    v = Array(DI_NORTHWEST, DI_NORTHEAST, DI_SOUTHWEST, DI_SOUTHEAST)
    For i = 0 To 3: diags(i) = v(i): Next
    v = Array(FRACUNIT, 47000, 0, -47000, -FRACUNIT, -47000, 0, 47000)
    For i = 0 To 7: xspeed(i) = v(i): Next
    v = Array(0, 47000, FRACUNIT, 47000, 0, -47000, -FRACUNIT, -47000)
    For i = 0 To 7: yspeed(i) = v(i): Next
    ReDim braintargets(0 To 31)
    tablesInit = True
End Sub

Private Function P_SubRandom() As Long
    Dim r As Long
    r = P_Random()
    P_SubRandom = r - P_Random()
End Function

' Mobj action dispatcher (C: st->action.acp1 (mobj)), called by P_SetMobjState.
Public Sub P_CallMobjAction(ByVal action As Long, ByVal mo As Long)
    If Not tablesInit Then InitTables
    Select Case action
        Case AF_Look: A_Look mo
        Case AF_Chase: A_Chase mo
        Case AF_FaceTarget: A_FaceTarget mo
        Case AF_PosAttack: A_PosAttack mo
        Case AF_Scream: A_Scream mo
        Case AF_XScream: S_StartSound mo, sfx_slop
        Case AF_Pain: A_Pain mo
        Case AF_Fall: A_Fall mo
        Case AF_Explode: P_RadiusAttack mo, mobjs(mo).target, 128
        Case AF_SPosAttack: A_SPosAttack mo
        Case AF_CPosAttack: A_CPosAttack mo
        Case AF_CPosRefire: A_CPosRefire mo
        Case AF_SpidRefire: A_SpidRefire mo
        Case AF_BspiAttack: A_BspiAttack mo
        Case AF_TroopAttack: A_TroopAttack mo
        Case AF_SargAttack: A_SargAttack mo
        Case AF_HeadAttack: A_HeadAttack mo
        Case AF_CyberAttack: A_CyberAttack mo
        Case AF_BruisAttack: A_BruisAttack mo
        Case AF_SkelMissile: A_SkelMissile mo
        Case AF_Tracer: A_Tracer mo
        Case AF_SkelWhoosh: A_SkelWhoosh mo
        Case AF_SkelFist: A_SkelFist mo
        Case AF_VileChase: A_VileChase mo
        Case AF_VileStart: S_StartSound mo, sfx_vilatk
        Case AF_VileTarget: A_VileTarget mo
        Case AF_VileAttack: A_VileAttack mo
        Case AF_StartFire
            S_StartSound mo, sfx_flamst
            A_Fire mo
        Case AF_Fire: A_Fire mo
        Case AF_FireCrackle
            S_StartSound mo, sfx_flame
            A_Fire mo
        Case AF_FatRaise
            A_FaceTarget mo
            S_StartSound mo, sfx_manatk
        Case AF_FatAttack1: A_FatAttack1 mo
        Case AF_FatAttack2: A_FatAttack2 mo
        Case AF_FatAttack3: A_FatAttack3 mo
        Case AF_SkullAttack: A_SkullAttack mo
        Case AF_PainAttack: A_PainAttack mo
        Case AF_PainDie: A_PainDie mo
        Case AF_BossDeath: A_BossDeath mo
        Case AF_Hoof
            S_StartSound mo, sfx_hoof
            A_Chase mo
        Case AF_Metal
            S_StartSound mo, sfx_metal
            A_Chase mo
        Case AF_BabyMetal
            S_StartSound mo, sfx_bspwlk
            A_Chase mo
        Case AF_BrainAwake: A_BrainAwake mo
        Case AF_BrainPain: S_StartSound 0, sfx_bospn
        Case AF_BrainScream: A_BrainScream mo
        Case AF_BrainExplode: A_BrainExplode mo
        Case AF_BrainDie: G_ExitLevel
        Case AF_BrainSpit: A_BrainSpit mo
        Case AF_SpawnSound
            S_StartSound mo, sfx_boscub
            A_SpawnFly mo
        Case AF_SpawnFly: A_SpawnFly mo
        Case AF_PlayerScream: A_PlayerScream mo
        Case AF_KeenDie: A_KeenDie mo
        Case AF_BFGSpray: A_BFGSpray mo
    End Select
End Sub

' ENEMY THINKING
' Enemies are allways spawned with targetplayer = -1, threshold = 0.  Most
' monsters are spawned unaware of all players, but some can be made preaware

' Called by P_NoiseAlert.  Recursively traverse adjacent sectors, sound
' blocking lines cut off traversal.
Private Sub P_RecursiveSound(ByVal sec As Long, ByVal soundblocks As Long)
    Dim i As Long, check As Long, other As Long

    ' wake up all monsters in this sector
    If sectors(sec).validcount = validcount And sectors(sec).soundtraversed <= soundblocks + 1 Then
        Exit Sub                                 ' already flooded
    End If

    sectors(sec).validcount = validcount
    sectors(sec).soundtraversed = soundblocks + 1
    sectors(sec).soundtarget = soundtarget

    For i = 0 To sectors(sec).linecount - 1
        check = linebuffer(sectors(sec).lines + i)
        If (lines(check).flags And ML_TWOSIDED) = 0 Then GoTo nextline

        P_LineOpening check

        If openrange <= 0 Then GoTo nextline     ' closed door

        If sides(lines(check).sidenum(0)).sector = sec Then
            other = sides(lines(check).sidenum(1)).sector
        Else
            other = sides(lines(check).sidenum(0)).sector
        End If

        If (lines(check).flags And ML_SOUNDBLOCK) <> 0 Then
            If soundblocks = 0 Then P_RecursiveSound other, 1
        Else
            P_RecursiveSound other, soundblocks
        End If
nextline:
    Next
End Sub

' P_NoiseAlert - if a monster yells at a player, it will alert other
' monsters to the player.
Public Sub P_NoiseAlert(ByVal target As Long, ByVal emmiter As Long)
    soundtarget = target
    validcount = validcount + 1
    P_RecursiveSound subsectors(mobjs(emmiter).subsector).sector, 0
End Sub

' P_CheckMeleeRange
Private Function P_CheckMeleeRange(ByVal actor As Long) As Boolean
    Dim pl As Long, dist As Long

    P_CheckMeleeRange = False
    If mobjs(actor).target = 0 Then Exit Function

    pl = mobjs(actor).target
    dist = P_AproxDistance(mobjs(pl).x - mobjs(actor).x, mobjs(pl).y - mobjs(actor).y)

    If dist >= MELEERANGE - 20 * FRACUNIT + mobjinfo(mobjs(pl).info).radius Then Exit Function

    If Not P_CheckSight(actor, mobjs(actor).target) Then Exit Function

    P_CheckMeleeRange = True
End Function

' P_CheckMissileRange
Private Function P_CheckMissileRange(ByVal actor As Long) As Boolean
    Dim dist As Long, tgt As Long

    P_CheckMissileRange = False
    tgt = mobjs(actor).target

    If Not P_CheckSight(actor, tgt) Then Exit Function

    If (mobjs(actor).flags And MF_JUSTHIT) <> 0 Then
        ' the target just hit the enemy, so fight back!
        mobjs(actor).flags = mobjs(actor).flags And (Not MF_JUSTHIT)
        P_CheckMissileRange = True
        Exit Function
    End If

    If mobjs(actor).reactiontime <> 0 Then Exit Function ' do not attack yet

    ' OPTIMIZE: get this from a global checksight
    dist = P_AproxDistance(mobjs(actor).x - mobjs(tgt).x, mobjs(actor).y - mobjs(tgt).y) - 64 * FRACUNIT

    If mobjinfo(mobjs(actor).info).meleestate = 0 Then dist = dist - 128 * FRACUNIT ' no melee attack, so fire more

    dist = SAR(dist, 16)

    If mobjs(actor).mtype = MT_VILE Then
        If dist > 14 * 64 Then Exit Function      ' too far away
    End If

    If mobjs(actor).mtype = MT_UNDEAD Then
        If dist < 196 Then Exit Function          ' close for fist attack
        dist = SAR(dist, 1)
    End If

    If mobjs(actor).mtype = MT_CYBORG Or mobjs(actor).mtype = MT_SPIDER Or mobjs(actor).mtype = MT_SKULL Then
        dist = SAR(dist, 1)
    End If

    If dist > 200 Then dist = 200

    If mobjs(actor).mtype = MT_CYBORG And dist > 160 Then dist = 160

    If P_Random() < dist Then Exit Function

    P_CheckMissileRange = True
End Function

' P_Move - move in the current direction, returns false if the move is blocked.
Private Function P_Move(ByVal actor As Long) As Boolean
    Dim tryx As Long, tryy As Long, ld As Long, try_ok As Boolean, good As Boolean, speed As Long

    If mobjs(actor).movedir = DI_NODIR Then
        P_Move = False
        Exit Function
    End If

    If mobjs(actor).movedir < 0 Or mobjs(actor).movedir >= 8 Then I_Error "Weird actor->movedir!"

    speed = mobjinfo(mobjs(actor).info).speed
    tryx = mobjs(actor).x + speed * xspeed(mobjs(actor).movedir)
    tryy = mobjs(actor).y + speed * yspeed(mobjs(actor).movedir)

    try_ok = P_TryMove(actor, tryx, tryy)

    If Not try_ok Then
        ' open any specials
        If (mobjs(actor).flags And MF_FLOAT) <> 0 And floatok Then
            ' must adjust height
            If mobjs(actor).z < tmfloorz Then
                mobjs(actor).z = mobjs(actor).z + FLOATSPEED
            Else
                mobjs(actor).z = mobjs(actor).z - FLOATSPEED
            End If

            mobjs(actor).flags = mobjs(actor).flags Or MF_INFLOAT
            P_Move = True
            Exit Function
        End If

        If numspechit = 0 Then
            P_Move = False
            Exit Function
        End If

        mobjs(actor).movedir = DI_NODIR
        good = False
        Do While numspechit > 0
            numspechit = numspechit - 1
            ld = spechit(numspechit)
            ' if the special is not a door that can be opened, return false
            If P_UseSpecialLine(actor, ld, 0) Then good = True
        Loop
        P_Move = good
        Exit Function
    Else
        mobjs(actor).flags = mobjs(actor).flags And (Not MF_INFLOAT)
    End If

    If (mobjs(actor).flags And MF_FLOAT) = 0 Then mobjs(actor).z = mobjs(actor).floorz
    P_Move = True
End Function

' TryWalk - attempts to move actor on in its current (ob->moveangle)
' direction.  If blocked by either a wall or an actor returns FALSE.  If move
' is either clear or blocked only by a door, returns TRUE and sets...  If a
' door is in the way, an OpenDoor call is made to start it opening.
Private Function P_TryWalk(ByVal actor As Long) As Boolean
    If Not P_Move(actor) Then
        P_TryWalk = False
        Exit Function
    End If

    mobjs(actor).movecount = P_Random() And 15
    P_TryWalk = True
End Function

Private Sub P_NewChaseDir(ByVal actor As Long)
    Dim deltax As Long, deltay As Long, d(0 To 2) As Long, tdir As Long
    Dim olddir As Long, turnaround As Long, tgt As Long, idx As Long

    tgt = mobjs(actor).target
    If tgt = 0 Then I_Error "P_NewChaseDir: called with no target"

    olddir = mobjs(actor).movedir
    turnaround = opposite(olddir)

    deltax = mobjs(tgt).x - mobjs(actor).x
    deltay = mobjs(tgt).y - mobjs(actor).y

    If deltax > 10 * FRACUNIT Then
        d(1) = DI_EAST
    ElseIf deltax < -10 * FRACUNIT Then
        d(1) = DI_WEST
    Else
        d(1) = DI_NODIR
    End If

    If deltay < -10 * FRACUNIT Then
        d(2) = DI_SOUTH
    ElseIf deltay > 10 * FRACUNIT Then
        d(2) = DI_NORTH
    Else
        d(2) = DI_NODIR
    End If

    ' try direct route
    If d(1) <> DI_NODIR And d(2) <> DI_NODIR Then
        idx = 0
        If deltay < 0 Then idx = 2
        If deltax > 0 Then idx = idx + 1
        mobjs(actor).movedir = diags(idx)
        If mobjs(actor).movedir <> turnaround Then
            If P_TryWalk(actor) Then Exit Sub
        End If
    End If

    ' try other directions
    If P_Random() > 200 Or Abs(deltay) > Abs(deltax) Then
        tdir = d(1)
        d(1) = d(2)
        d(2) = tdir
    End If

    If d(1) = turnaround Then d(1) = DI_NODIR
    If d(2) = turnaround Then d(2) = DI_NODIR

    If d(1) <> DI_NODIR Then
        mobjs(actor).movedir = d(1)
        If P_TryWalk(actor) Then Exit Sub        ' either moved forward or attacked
    End If

    If d(2) <> DI_NODIR Then
        mobjs(actor).movedir = d(2)
        If P_TryWalk(actor) Then Exit Sub
    End If

    ' there is no direct path to the player, so pick another direction.
    If olddir <> DI_NODIR Then
        mobjs(actor).movedir = olddir
        If P_TryWalk(actor) Then Exit Sub
    End If

    ' randomly determine direction of search
    If (P_Random() And 1) <> 0 Then
        For tdir = DI_EAST To DI_SOUTHEAST
            If tdir <> turnaround Then
                mobjs(actor).movedir = tdir
                If P_TryWalk(actor) Then Exit Sub
            End If
        Next
    Else
        For tdir = DI_SOUTHEAST To DI_EAST Step -1
            If tdir <> turnaround Then
                mobjs(actor).movedir = tdir
                If P_TryWalk(actor) Then Exit Sub
            End If
        Next
    End If

    If turnaround <> DI_NODIR Then
        mobjs(actor).movedir = turnaround
        If P_TryWalk(actor) Then Exit Sub
    End If

    mobjs(actor).movedir = DI_NODIR              ' can not move
End Sub

' P_LookForPlayers - if allaround is false, only look 180 degrees in front.
' Returns true if a player is targeted.
Private Function P_LookForPlayers(ByVal actor As Long, ByVal allaround As Boolean) As Boolean
    Dim c As Long, stopl As Long, player As Long, an As Long, dist As Long, pmo As Long

    c = 0
    stopl = (mobjs(actor).lastlook - 1) And 3

    Do
        If Not playeringame(mobjs(actor).lastlook) Then GoTo nextlook

        ' C: if (c++ == 2 || actor->lastlook == stop)
        If c = 2 Then
            c = c + 1
            P_LookForPlayers = False             ' done looking
            Exit Function
        End If
        c = c + 1
        If mobjs(actor).lastlook = stopl Then
            P_LookForPlayers = False             ' done looking
            Exit Function
        End If

        player = mobjs(actor).lastlook

        If players(player).health <= 0 Then GoTo nextlook ' dead

        pmo = players(player).mo
        If Not P_CheckSight(actor, pmo) Then GoTo nextlook ' out of sight

        If Not allaround Then
            an = R_PointToAngle2(mobjs(actor).x, mobjs(actor).y, mobjs(pmo).x, mobjs(pmo).y) - mobjs(actor).angle

            If UGT(an, ANG90) And ULT(an, ANG270) Then
                dist = P_AproxDistance(mobjs(pmo).x - mobjs(actor).x, mobjs(pmo).y - mobjs(actor).y)
                ' if real close, react anyway
                If dist > MELEERANGE Then GoTo nextlook ' behind back
            End If
        End If

        mobjs(actor).target = pmo
        P_LookForPlayers = True
        Exit Function
nextlook:
        mobjs(actor).lastlook = (mobjs(actor).lastlook + 1) And 3
    Loop
End Function

' A_KeenDie - DOOM II special, map 32.  Uses special tag 666.
Private Sub A_KeenDie(ByVal mo As Long)
    Dim th As Long

    A_Fall mo

    ' scan the remaining thinkers to see if all Keens are dead
    th = thinkers(0).nxt
    Do While th <> 0
        If thinkers(th).func = TF_MOBJ Then
            If th <> mo And mobjs(th).mtype = mobjs(mo).mtype And mobjs(th).health > 0 Then
                ' other Keen not dead
                Exit Sub
            End If
        End If
        th = thinkers(th).nxt
    Loop

    lines(numlines).tag = 666                    ' the scratch "junk" line
    EV_DoDoor numlines, vld_open
End Sub

' ACTION ROUTINES

' A_Look - stay in state until a player is sighted.
Private Sub A_Look(ByVal actor As Long)
    Dim targ As Long, sound As Long, seen As Boolean, ss As Long

    mobjs(actor).threshold = 0                   ' any shot will wake up
    targ = sectors(subsectors(mobjs(actor).subsector).sector).soundtarget

    seen = False
    If targ <> 0 Then
        If (mobjs(targ).flags And MF_SHOOTABLE) <> 0 Then
            mobjs(actor).target = targ

            If (mobjs(actor).flags And MF_AMBUSH) <> 0 Then
                If P_CheckSight(actor, mobjs(actor).target) Then seen = True
            Else
                seen = True
            End If
        End If
    End If

    If Not seen Then
        If Not P_LookForPlayers(actor, False) Then Exit Sub
    End If

    ' go into chase state
    ss = mobjinfo(mobjs(actor).info).seesound
    If ss <> 0 Then
        Select Case ss
            Case sfx_posit1, sfx_posit2, sfx_posit3
                sound = sfx_posit1 + P_Random() Mod 3
            Case sfx_bgsit1, sfx_bgsit2
                sound = sfx_bgsit1 + P_Random() Mod 2
            Case Else
                sound = ss
        End Select

        If mobjs(actor).mtype = MT_SPIDER Or mobjs(actor).mtype = MT_CYBORG Then
            ' full volume
            S_StartSound 0, sound
        Else
            S_StartSound actor, sound
        End If
    End If

    P_SetMobjState actor, mobjinfo(mobjs(actor).info).seestate
End Sub

' A_Chase - actor has a melee attack, so it tries to close as fast as
' possible
Private Sub A_Chase(ByVal actor As Long)
    Dim delta As Long, tgt As Long, info As Long

    info = mobjs(actor).info

    If mobjs(actor).reactiontime <> 0 Then mobjs(actor).reactiontime = mobjs(actor).reactiontime - 1

    ' modify target threshold
    If mobjs(actor).threshold <> 0 Then
        tgt = mobjs(actor).target
        If tgt = 0 Then
            mobjs(actor).threshold = 0
        ElseIf mobjs(tgt).health <= 0 Then
            mobjs(actor).threshold = 0
        Else
            mobjs(actor).threshold = mobjs(actor).threshold - 1
        End If
    End If

    ' turn towards movement direction if not there yet
    If mobjs(actor).movedir < 8 Then
        mobjs(actor).angle = mobjs(actor).angle And &HE0000000
        delta = mobjs(actor).angle - (mobjs(actor).movedir * &H20000000)

        If delta > 0 Then
            mobjs(actor).angle = mobjs(actor).angle - ANG90 \ 2
        ElseIf delta < 0 Then
            mobjs(actor).angle = mobjs(actor).angle + ANG90 \ 2
        End If
    End If

    tgt = mobjs(actor).target
    If tgt = 0 Then GoTo lookagain
    If (mobjs(tgt).flags And MF_SHOOTABLE) = 0 Then GoTo lookagain
    GoTo havetarget

lookagain:
    ' look for a new target
    If P_LookForPlayers(actor, True) Then Exit Sub ' got a new target

    P_SetMobjState actor, mobjinfo(info).spawnstate
    Exit Sub

havetarget:
    ' do not attack twice in a row
    If (mobjs(actor).flags And MF_JUSTATTACKED) <> 0 Then
        mobjs(actor).flags = mobjs(actor).flags And (Not MF_JUSTATTACKED)
        If gameskill <> sk_nightmare And Not fastparm Then P_NewChaseDir actor
        Exit Sub
    End If

    ' check for melee attack
    If mobjinfo(info).meleestate <> 0 Then
        If P_CheckMeleeRange(actor) Then
            If mobjinfo(info).attacksound <> 0 Then S_StartSound actor, mobjinfo(info).attacksound

            P_SetMobjState actor, mobjinfo(info).meleestate
            Exit Sub
        End If
    End If

    ' check for missile attack
    If mobjinfo(info).missilestate <> 0 Then
        If gameskill < sk_nightmare And Not fastparm And mobjs(actor).movecount <> 0 Then GoTo nomissile

        If Not P_CheckMissileRange(actor) Then GoTo nomissile

        P_SetMobjState actor, mobjinfo(info).missilestate
        mobjs(actor).flags = mobjs(actor).flags Or MF_JUSTATTACKED
        Exit Sub
    End If

nomissile:
    ' possibly choose another target
    If netgame And mobjs(actor).threshold = 0 Then
        If Not P_CheckSight(actor, mobjs(actor).target) Then
            If P_LookForPlayers(actor, True) Then Exit Sub ' got a new target
        End If
    End If

    ' chase towards player
    mobjs(actor).movecount = mobjs(actor).movecount - 1
    If mobjs(actor).movecount < 0 Then
        P_NewChaseDir actor
    ElseIf Not P_Move(actor) Then
        P_NewChaseDir actor
    End If

    ' make active sound
    If mobjinfo(info).activesound <> 0 Then
        If P_Random() < 3 Then S_StartSound actor, mobjinfo(info).activesound
    End If
End Sub

' A_FaceTarget
Private Sub A_FaceTarget(ByVal actor As Long)
    Dim tgt As Long
    tgt = mobjs(actor).target
    If tgt = 0 Then Exit Sub

    mobjs(actor).flags = mobjs(actor).flags And (Not MF_AMBUSH)

    mobjs(actor).angle = R_PointToAngle2(mobjs(actor).x, mobjs(actor).y, mobjs(tgt).x, mobjs(tgt).y)

    If (mobjs(tgt).flags And MF_SHADOW) <> 0 Then
        mobjs(actor).angle = mobjs(actor).angle + P_SubRandom() * &H200000
    End If
End Sub

' A_PosAttack
Private Sub A_PosAttack(ByVal actor As Long)
    Dim angle As Long, damage As Long, slope As Long

    If mobjs(actor).target = 0 Then Exit Sub

    A_FaceTarget actor
    angle = mobjs(actor).angle
    slope = P_AimLineAttack(actor, angle, MISSILERANGE)

    S_StartSound actor, sfx_pistol
    angle = angle + P_SubRandom() * &H100000
    damage = ((P_Random() Mod 5) + 1) * 3
    P_LineAttack actor, angle, MISSILERANGE, slope, damage
End Sub

Private Sub A_SPosAttack(ByVal actor As Long)
    Dim i As Long, angle As Long, bangle As Long, damage As Long, slope As Long

    If mobjs(actor).target = 0 Then Exit Sub

    S_StartSound actor, sfx_shotgn
    A_FaceTarget actor
    bangle = mobjs(actor).angle
    slope = P_AimLineAttack(actor, bangle, MISSILERANGE)

    For i = 0 To 2
        angle = bangle + P_SubRandom() * &H100000
        damage = ((P_Random() Mod 5) + 1) * 3
        P_LineAttack actor, angle, MISSILERANGE, slope, damage
    Next
End Sub

Private Sub A_CPosAttack(ByVal actor As Long)
    Dim angle As Long, bangle As Long, damage As Long, slope As Long

    If mobjs(actor).target = 0 Then Exit Sub

    S_StartSound actor, sfx_shotgn
    A_FaceTarget actor
    bangle = mobjs(actor).angle
    slope = P_AimLineAttack(actor, bangle, MISSILERANGE)

    angle = bangle + P_SubRandom() * &H100000
    damage = ((P_Random() Mod 5) + 1) * 3
    P_LineAttack actor, angle, MISSILERANGE, slope, damage
End Sub

Private Sub A_CPosRefire(ByVal actor As Long)
    Dim tgt As Long

    ' keep firing unless target got out of sight
    A_FaceTarget actor

    If P_Random() < 40 Then Exit Sub

    tgt = mobjs(actor).target
    If tgt = 0 Then
        P_SetMobjState actor, mobjinfo(mobjs(actor).info).seestate
    ElseIf mobjs(tgt).health <= 0 Then
        P_SetMobjState actor, mobjinfo(mobjs(actor).info).seestate
    ElseIf Not P_CheckSight(actor, tgt) Then
        P_SetMobjState actor, mobjinfo(mobjs(actor).info).seestate
    End If
End Sub

Private Sub A_SpidRefire(ByVal actor As Long)
    Dim tgt As Long

    ' keep firing unless target got out of sight
    A_FaceTarget actor

    If P_Random() < 10 Then Exit Sub

    tgt = mobjs(actor).target
    If tgt = 0 Then
        P_SetMobjState actor, mobjinfo(mobjs(actor).info).seestate
    ElseIf mobjs(tgt).health <= 0 Then
        P_SetMobjState actor, mobjinfo(mobjs(actor).info).seestate
    ElseIf Not P_CheckSight(actor, tgt) Then
        P_SetMobjState actor, mobjinfo(mobjs(actor).info).seestate
    End If
End Sub

Private Sub A_BspiAttack(ByVal actor As Long)
    If mobjs(actor).target = 0 Then Exit Sub

    A_FaceTarget actor

    ' launch a missile
    P_SpawnMissile actor, mobjs(actor).target, MT_ARACHPLAZ
End Sub

' A_TroopAttack
Private Sub A_TroopAttack(ByVal actor As Long)
    Dim damage As Long

    If mobjs(actor).target = 0 Then Exit Sub

    A_FaceTarget actor
    If P_CheckMeleeRange(actor) Then
        S_StartSound actor, sfx_claw
        damage = ((P_Random() Mod 8) + 1) * 3
        P_DamageMobj mobjs(actor).target, actor, actor, damage
        Exit Sub
    End If

    ' launch a missile
    P_SpawnMissile actor, mobjs(actor).target, MT_TROOPSHOT
End Sub

Private Sub A_SargAttack(ByVal actor As Long)
    Dim damage As Long

    If mobjs(actor).target = 0 Then Exit Sub

    A_FaceTarget actor
    If P_CheckMeleeRange(actor) Then
        damage = ((P_Random() Mod 10) + 1) * 4
        P_DamageMobj mobjs(actor).target, actor, actor, damage
    End If
End Sub

Private Sub A_HeadAttack(ByVal actor As Long)
    Dim damage As Long

    If mobjs(actor).target = 0 Then Exit Sub

    A_FaceTarget actor
    If P_CheckMeleeRange(actor) Then
        damage = ((P_Random() Mod 6) + 1) * 10
        P_DamageMobj mobjs(actor).target, actor, actor, damage
        Exit Sub
    End If

    ' launch a missile
    P_SpawnMissile actor, mobjs(actor).target, MT_HEADSHOT
End Sub

Private Sub A_CyberAttack(ByVal actor As Long)
    If mobjs(actor).target = 0 Then Exit Sub

    A_FaceTarget actor
    P_SpawnMissile actor, mobjs(actor).target, MT_ROCKET
End Sub

Private Sub A_BruisAttack(ByVal actor As Long)
    Dim damage As Long

    If mobjs(actor).target = 0 Then Exit Sub

    If P_CheckMeleeRange(actor) Then
        S_StartSound actor, sfx_claw
        damage = ((P_Random() Mod 8) + 1) * 10
        P_DamageMobj mobjs(actor).target, actor, actor, damage
        Exit Sub
    End If

    ' launch a missile
    P_SpawnMissile actor, mobjs(actor).target, MT_BRUISERSHOT
End Sub

' A_SkelMissile
Private Sub A_SkelMissile(ByVal actor As Long)
    Dim mo As Long

    If mobjs(actor).target = 0 Then Exit Sub

    A_FaceTarget actor
    mobjs(actor).z = mobjs(actor).z + 16 * FRACUNIT ' so missile spawns higher
    mo = P_SpawnMissile(actor, mobjs(actor).target, MT_TRACER)
    mobjs(actor).z = mobjs(actor).z - 16 * FRACUNIT ' back to normal

    mobjs(mo).x = mobjs(mo).x + mobjs(mo).momx
    mobjs(mo).y = mobjs(mo).y + mobjs(mo).momy
    mobjs(mo).tracer = mobjs(actor).target
End Sub

Private Sub A_Tracer(ByVal actor As Long)
    Dim exact As Long, dist As Long, slope As Long, dest As Long, th As Long, speed As Long

    If (gametic And 3) <> 0 Then Exit Sub

    ' spawn a puff of smoke behind the rocket
    P_SpawnPuff mobjs(actor).x, mobjs(actor).y, mobjs(actor).z

    th = P_SpawnMobj(mobjs(actor).x - mobjs(actor).momx, mobjs(actor).y - mobjs(actor).momy, mobjs(actor).z, MT_SMOKE)

    mobjs(th).momz = FRACUNIT
    mobjs(th).tics = mobjs(th).tics - (P_Random() And 3)
    If mobjs(th).tics < 1 Then mobjs(th).tics = 1

    ' adjust direction
    dest = mobjs(actor).tracer

    If dest = 0 Then Exit Sub
    If mobjs(dest).health <= 0 Then Exit Sub

    ' change angle
    exact = R_PointToAngle2(mobjs(actor).x, mobjs(actor).y, mobjs(dest).x, mobjs(dest).y)

    If exact <> mobjs(actor).angle Then
        If UGT(exact - mobjs(actor).angle, &H80000000) Then
            mobjs(actor).angle = mobjs(actor).angle - TRACEANGLE
            If ULT(exact - mobjs(actor).angle, &H80000000) Then mobjs(actor).angle = exact
        Else
            mobjs(actor).angle = mobjs(actor).angle + TRACEANGLE
            If UGT(exact - mobjs(actor).angle, &H80000000) Then mobjs(actor).angle = exact
        End If
    End If

    exact = FINEANGLE(mobjs(actor).angle)
    speed = mobjinfo(mobjs(actor).info).speed
    mobjs(actor).momx = FixedMul(speed, finecosine(exact))
    mobjs(actor).momy = FixedMul(speed, finesine(exact))

    ' change slope
    dist = P_AproxDistance(mobjs(dest).x - mobjs(actor).x, mobjs(dest).y - mobjs(actor).y)

    dist = dist \ speed

    If dist < 1 Then dist = 1
    slope = (mobjs(dest).z + 40 * FRACUNIT - mobjs(actor).z) \ dist

    If slope < mobjs(actor).momz Then
        mobjs(actor).momz = mobjs(actor).momz - FRACUNIT \ 8
    Else
        mobjs(actor).momz = mobjs(actor).momz + FRACUNIT \ 8
    End If
End Sub

Private Sub A_SkelWhoosh(ByVal actor As Long)
    If mobjs(actor).target = 0 Then Exit Sub
    A_FaceTarget actor
    S_StartSound actor, sfx_skeswg
End Sub

Private Sub A_SkelFist(ByVal actor As Long)
    Dim damage As Long

    If mobjs(actor).target = 0 Then Exit Sub

    A_FaceTarget actor

    If P_CheckMeleeRange(actor) Then
        damage = ((P_Random() Mod 10) + 1) * 6
        S_StartSound actor, sfx_skepch
        P_DamageMobj mobjs(actor).target, actor, actor, damage
    End If
End Sub

' PIT_VileCheck - detect a corpse that could be raised.
Public Function PIT_VileCheck(ByVal thing As Long) As Boolean
    Dim maxdist As Long, check As Boolean

    PIT_VileCheck = True

    If (mobjs(thing).flags And MF_CORPSE) = 0 Then Exit Function ' not a monster

    If mobjs(thing).tics <> -1 Then Exit Function ' not lying still yet

    If mobjinfo(mobjs(thing).info).raisestate = S_NULL Then Exit Function ' monster doesn't have a raise state

    maxdist = mobjinfo(mobjs(thing).info).radius + mobjinfo(MT_VILE).radius

    If Abs(mobjs(thing).x - viletryx) > maxdist Or Abs(mobjs(thing).y - viletryy) > maxdist Then
        Exit Function                            ' not actually touching
    End If

    corpsehit = thing
    mobjs(corpsehit).momx = 0
    mobjs(corpsehit).momy = 0
    mobjs(corpsehit).height = mobjs(corpsehit).height * 4
    check = P_CheckPosition(corpsehit, mobjs(corpsehit).x, mobjs(corpsehit).y)
    mobjs(corpsehit).height = SAR(mobjs(corpsehit).height, 2)

    If Not check Then Exit Function              ' doesn't fit here

    PIT_VileCheck = False                        ' got one, so stop checking
End Function

' A_VileChase - check for ressurecting a body
Private Sub A_VileChase(ByVal actor As Long)
    Dim xl As Long, xh As Long, yl As Long, yh As Long, bx As Long, by As Long
    Dim info As Long, temp As Long, speed As Long

    If mobjs(actor).movedir <> DI_NODIR Then
        ' check for corpses to raise
        speed = mobjinfo(mobjs(actor).info).speed
        viletryx = mobjs(actor).x + speed * xspeed(mobjs(actor).movedir)
        viletryy = mobjs(actor).y + speed * yspeed(mobjs(actor).movedir)

        xl = SAR(viletryx - bmaporgx - MAXRADIUS * 2, MAPBLOCKSHIFT)
        xh = SAR(viletryx - bmaporgx + MAXRADIUS * 2, MAPBLOCKSHIFT)
        yl = SAR(viletryy - bmaporgy - MAXRADIUS * 2, MAPBLOCKSHIFT)
        yh = SAR(viletryy - bmaporgy + MAXRADIUS * 2, MAPBLOCKSHIFT)

        vileobj = actor
        For bx = xl To xh
            For by = yl To yh
                ' Call PIT_VileCheck to check whether object is a corpse that
                ' can be raised.
                If Not P_BlockThingsIterator(bx, by, ITER_VILECHECK) Then
                    ' got one!
                    temp = mobjs(actor).target
                    mobjs(actor).target = corpsehit
                    A_FaceTarget actor
                    mobjs(actor).target = temp

                    P_SetMobjState actor, S_VILE_HEAL1
                    S_StartSound corpsehit, sfx_slop
                    info = mobjs(corpsehit).info

                    P_SetMobjState corpsehit, mobjinfo(info).raisestate
                    mobjs(corpsehit).height = mobjs(corpsehit).height * 4
                    mobjs(corpsehit).flags = mobjinfo(info).flags
                    mobjs(corpsehit).health = mobjinfo(info).spawnhealth
                    mobjs(corpsehit).target = 0

                    Exit Sub
                End If
            Next
        Next
    End If

    ' Return to normal attack.
    A_Chase actor
End Sub

' A_Fire - keep fire in front of player unless out of sight
Private Sub A_Fire(ByVal actor As Long)
    Dim dest As Long, an As Long

    dest = mobjs(actor).tracer
    If dest = 0 Then Exit Sub

    ' don't move it if the vile lost sight
    If Not P_CheckSight(mobjs(actor).target, dest) Then Exit Sub

    an = FINEANGLE(mobjs(dest).angle)
    If checklinks Then
        dbgUnsetBy = "A_Fire(" & actor & ") thinker " & currentThinker & " type " & mobjs(currentThinker).mtype & " state " & mobjs(currentThinker).state
        If thinkers(actor).func <> TF_MOBJ Then Err.Raise vbObjectError + 2, , "A_Fire on dead slot " & actor & " (type " & mobjs(actor).mtype & ") by " & dbgUnsetBy & " at gametic " & gametic
    End If

    P_UnsetThingPosition actor
    mobjs(actor).x = mobjs(dest).x + FixedMul(24 * FRACUNIT, finecosine(an))
    mobjs(actor).y = mobjs(dest).y + FixedMul(24 * FRACUNIT, finesine(an))
    mobjs(actor).z = mobjs(dest).z
    P_SetThingPosition actor
End Sub

' A_VileTarget - spawn the hellfire
Private Sub A_VileTarget(ByVal actor As Long)
    Dim fog As Long, tgt As Long

    tgt = mobjs(actor).target
    If tgt = 0 Then Exit Sub

    A_FaceTarget actor

    ' (sic: the C code passes target->x twice)
    fog = P_SpawnMobj(mobjs(tgt).x, mobjs(tgt).x, mobjs(tgt).z, MT_FIRE)

    mobjs(actor).tracer = fog
    mobjs(fog).target = actor
    mobjs(fog).tracer = mobjs(actor).target
    A_Fire fog
End Sub

' A_VileAttack
Private Sub A_VileAttack(ByVal actor As Long)
    Dim fire As Long, an As Long, tgt As Long

    tgt = mobjs(actor).target
    If tgt = 0 Then Exit Sub

    A_FaceTarget actor

    If Not P_CheckSight(actor, tgt) Then Exit Sub

    S_StartSound actor, sfx_barexp
    P_DamageMobj tgt, actor, actor, 20
    mobjs(tgt).momz = 1000 * FRACUNIT \ mobjinfo(mobjs(tgt).info).mass

    an = FINEANGLE(mobjs(actor).angle)

    fire = mobjs(actor).tracer

    If fire = 0 Then Exit Sub

    ' move the fire between the vile and the player
    mobjs(fire).x = mobjs(tgt).x - FixedMul(24 * FRACUNIT, finecosine(an))
    mobjs(fire).y = mobjs(tgt).y - FixedMul(24 * FRACUNIT, finesine(an))
    P_RadiusAttack fire, actor, 70
End Sub

' Mancubus attack, firing three missiles (bruisers) in three different
' directions?  Doesn't look like it.
Private Sub FatAim(ByVal mo As Long)
    Dim an As Long, speed As Long
    an = FINEANGLE(mobjs(mo).angle)
    speed = mobjinfo(mobjs(mo).info).speed
    mobjs(mo).momx = FixedMul(speed, finecosine(an))
    mobjs(mo).momy = FixedMul(speed, finesine(an))
End Sub

Private Sub A_FatAttack1(ByVal actor As Long)
    Dim mo As Long

    A_FaceTarget actor
    ' Change direction  to ...
    mobjs(actor).angle = mobjs(actor).angle + FATSPREAD
    P_SpawnMissile actor, mobjs(actor).target, MT_FATSHOT

    mo = P_SpawnMissile(actor, mobjs(actor).target, MT_FATSHOT)
    mobjs(mo).angle = mobjs(mo).angle + FATSPREAD
    FatAim mo
End Sub

Private Sub A_FatAttack2(ByVal actor As Long)
    Dim mo As Long

    A_FaceTarget actor
    ' Now here choose opposite deviation.
    mobjs(actor).angle = mobjs(actor).angle - FATSPREAD
    P_SpawnMissile actor, mobjs(actor).target, MT_FATSHOT

    mo = P_SpawnMissile(actor, mobjs(actor).target, MT_FATSHOT)
    mobjs(mo).angle = mobjs(mo).angle - FATSPREAD * 2
    FatAim mo
End Sub

Private Sub A_FatAttack3(ByVal actor As Long)
    Dim mo As Long

    A_FaceTarget actor

    mo = P_SpawnMissile(actor, mobjs(actor).target, MT_FATSHOT)
    mobjs(mo).angle = mobjs(mo).angle - FATSPREAD \ 2
    FatAim mo

    mo = P_SpawnMissile(actor, mobjs(actor).target, MT_FATSHOT)
    mobjs(mo).angle = mobjs(mo).angle + FATSPREAD \ 2
    FatAim mo
End Sub

' SkullAttack - fly at the player like a missile.
Private Sub A_SkullAttack(ByVal actor As Long)
    Dim dest As Long, an As Long, dist As Long

    dest = mobjs(actor).target
    If dest = 0 Then Exit Sub

    mobjs(actor).flags = mobjs(actor).flags Or MF_SKULLFLY

    S_StartSound actor, mobjinfo(mobjs(actor).info).attacksound
    A_FaceTarget actor
    an = FINEANGLE(mobjs(actor).angle)
    mobjs(actor).momx = FixedMul(SKULLSPEED, finecosine(an))
    mobjs(actor).momy = FixedMul(SKULLSPEED, finesine(an))
    dist = P_AproxDistance(mobjs(dest).x - mobjs(actor).x, mobjs(dest).y - mobjs(actor).y)
    dist = dist \ SKULLSPEED

    If dist < 1 Then dist = 1
    mobjs(actor).momz = (mobjs(dest).z + SAR(mobjs(dest).height, 1) - mobjs(actor).z) \ dist
End Sub

' A_PainShootSkull - spawn a lost soul and launch it at the target
Private Sub A_PainShootSkull(ByVal actor As Long, ByVal angle As Long)
    Dim x As Long, y As Long, z As Long, newmobj As Long, an As Long
    Dim prestep As Long, count As Long, th As Long

    ' count total number of skull currently on the level
    count = 0

    th = thinkers(0).nxt
    Do While th <> 0
        If thinkers(th).func = TF_MOBJ Then
            If mobjs(th).mtype = MT_SKULL Then count = count + 1
        End If
        th = thinkers(th).nxt
    Loop

    ' if there are allready 20 skulls on the level, don't spit another one
    If count > 20 Then Exit Sub

    ' okay, there's playe for another one
    an = FINEANGLE(angle)

    prestep = 4 * FRACUNIT + 3 * (mobjinfo(mobjs(actor).info).radius + mobjinfo(MT_SKULL).radius) \ 2

    x = mobjs(actor).x + FixedMul(prestep, finecosine(an))
    y = mobjs(actor).y + FixedMul(prestep, finesine(an))
    z = mobjs(actor).z + 8 * FRACUNIT

    newmobj = P_SpawnMobj(x, y, z, MT_SKULL)

    ' Check for movements.
    If Not P_TryMove(newmobj, mobjs(newmobj).x, mobjs(newmobj).y) Then
        ' kill it immediately
        P_DamageMobj newmobj, actor, actor, 10000
        Exit Sub
    End If

    mobjs(newmobj).target = mobjs(actor).target
    A_SkullAttack newmobj
End Sub

' A_PainAttack - spawn a lost soul and launch it at the target
Private Sub A_PainAttack(ByVal actor As Long)
    If mobjs(actor).target = 0 Then Exit Sub

    A_FaceTarget actor
    A_PainShootSkull actor, mobjs(actor).angle
End Sub

Private Sub A_PainDie(ByVal actor As Long)
    A_Fall actor
    A_PainShootSkull actor, mobjs(actor).angle + ANG90
    A_PainShootSkull actor, mobjs(actor).angle + ANG180
    A_PainShootSkull actor, mobjs(actor).angle + ANG270
End Sub

Private Sub A_Scream(ByVal actor As Long)
    Dim sound As Long, ds As Long

    ds = mobjinfo(mobjs(actor).info).deathsound
    Select Case ds
        Case 0
            Exit Sub
        Case sfx_podth1, sfx_podth2, sfx_podth3
            sound = sfx_podth1 + P_Random() Mod 3
        Case sfx_bgdth1, sfx_bgdth2
            sound = sfx_bgdth1 + P_Random() Mod 2
        Case Else
            sound = ds
    End Select

    ' Check for bosses.
    If mobjs(actor).mtype = MT_SPIDER Or mobjs(actor).mtype = MT_CYBORG Then
        ' full volume
        S_StartSound 0, sound
    Else
        S_StartSound actor, sound
    End If
End Sub

Private Sub A_Pain(ByVal actor As Long)
    If mobjinfo(mobjs(actor).info).painsound <> 0 Then S_StartSound actor, mobjinfo(mobjs(actor).info).painsound
End Sub

Private Sub A_Fall(ByVal actor As Long)
    ' actor is on ground, it can be walked over
    mobjs(actor).flags = mobjs(actor).flags And (Not MF_SOLID)

    ' So change this if corpse objects are meant to be obstacles.
End Sub

' A_BossDeath - possibly trigger special effects if on first boss level
Private Sub A_BossDeath(ByVal mo As Long)
    Dim th As Long, i As Long, junk As Long, mt As Long

    mt = mobjs(mo).mtype
    If gamemode = commercial Then
        If gamemap <> 7 Then Exit Sub

        If mt <> MT_FATSO And mt <> MT_BABY Then Exit Sub
    Else
        Select Case gameepisode
            Case 1
                If gamemap <> 8 Then Exit Sub
                If mt <> MT_BRUISER Then Exit Sub
            Case 2
                If gamemap <> 8 Then Exit Sub
                If mt <> MT_CYBORG Then Exit Sub
            Case 3
                If gamemap <> 8 Then Exit Sub
                If mt <> MT_SPIDER Then Exit Sub
            Case 4
                Select Case gamemap
                    Case 6
                        If mt <> MT_CYBORG Then Exit Sub
                    Case 8
                        If mt <> MT_SPIDER Then Exit Sub
                    Case Else
                        Exit Sub
                End Select
            Case Else
                If gamemap <> 8 Then Exit Sub
        End Select
    End If

    ' make sure there is a player alive for victory
    For i = 0 To MAXPLAYERS - 1
        If playeringame(i) And players(i).health > 0 Then Exit For
    Next

    If i = MAXPLAYERS Then Exit Sub              ' no one left alive, so do not end game

    ' scan the remaining thinkers to see if all bosses are dead
    th = thinkers(0).nxt
    Do While th <> 0
        If thinkers(th).func = TF_MOBJ Then
            If th <> mo And mobjs(th).mtype = mt And mobjs(th).health > 0 Then
                ' other boss not dead
                Exit Sub
            End If
        End If
        th = thinkers(th).nxt
    Loop

    ' victory!
    junk = numlines                              ' the scratch "junk" line
    If gamemode = commercial Then
        If gamemap = 7 Then
            If mt = MT_FATSO Then
                lines(junk).tag = 666
                EV_DoFloor junk, lowerFloorToLowest
                Exit Sub
            End If

            If mt = MT_BABY Then
                lines(junk).tag = 667
                EV_DoFloor junk, raiseToTexture
                Exit Sub
            End If
        End If
    Else
        Select Case gameepisode
            Case 1
                lines(junk).tag = 666
                EV_DoFloor junk, lowerFloorToLowest
                Exit Sub
            Case 4
                Select Case gamemap
                    Case 6
                        lines(junk).tag = 666
                        EV_DoDoor junk, vld_blazeOpen
                        Exit Sub
                    Case 8
                        lines(junk).tag = 666
                        EV_DoFloor junk, lowerFloorToLowest
                        Exit Sub
                End Select
        End Select
    End If

    G_ExitLevel
End Sub

Private Sub A_BrainAwake(ByVal mo As Long)
    Dim th As Long

    ' find all the target spots
    numbraintargets = 0
    braintargeton = 0

    th = thinkers(0).nxt
    Do While th <> 0
        If thinkers(th).func = TF_MOBJ Then
            If mobjs(th).mtype = MT_BOSSTARGET Then
                If numbraintargets > UBound(braintargets) Then
                    ReDim Preserve braintargets(0 To numbraintargets * 2)
                End If
                braintargets(numbraintargets) = th
                numbraintargets = numbraintargets + 1
            End If
        End If
        th = thinkers(th).nxt
    Loop

    S_StartSound 0, sfx_bossit
End Sub

Private Sub A_BrainScream(ByVal mo As Long)
    Dim x As Long, y As Long, z As Long, th As Long

    x = mobjs(mo).x - 196 * FRACUNIT
    Do While x < mobjs(mo).x + 320 * FRACUNIT
        y = mobjs(mo).y - 320 * FRACUNIT
        z = 128 + P_Random() * 2 * FRACUNIT
        th = P_SpawnMobj(x, y, z, MT_ROCKET)
        mobjs(th).momz = P_Random() * 512

        P_SetMobjState th, S_BRAINEXPLODE1

        mobjs(th).tics = mobjs(th).tics - (P_Random() And 7)
        If mobjs(th).tics < 1 Then mobjs(th).tics = 1
        x = x + FRACUNIT * 8
    Loop

    S_StartSound 0, sfx_bosdth
End Sub

Private Sub A_BrainExplode(ByVal mo As Long)
    Dim x As Long, y As Long, z As Long, th As Long

    x = mobjs(mo).x + P_SubRandom() * 2048
    y = mobjs(mo).y
    z = 128 + P_Random() * 2 * FRACUNIT
    th = P_SpawnMobj(x, y, z, MT_ROCKET)
    mobjs(th).momz = P_Random() * 512

    P_SetMobjState th, S_BRAINEXPLODE1

    mobjs(th).tics = mobjs(th).tics - (P_Random() And 7)
    If mobjs(th).tics < 1 Then mobjs(th).tics = 1
End Sub

Private Sub A_BrainSpit(ByVal mo As Long)
    Dim targ As Long, newmobj As Long, t As Long

    brainEasy = brainEasy Xor 1
    If gameskill <= sk_easy And brainEasy = 0 Then Exit Sub

    If numbraintargets = 0 Then Exit Sub         ' (C would divide by zero)

    ' shoot a cube at current target
    targ = braintargets(braintargeton)
    braintargeton = (braintargeton + 1) Mod numbraintargets

    ' spawn brain missile
    newmobj = P_SpawnMissile(mo, targ, MT_SPAWNSHOT)
    mobjs(newmobj).target = targ
    t = states(mobjs(newmobj).state).tics
    If mobjs(newmobj).momy <> 0 And t <> 0 Then
        mobjs(newmobj).reactiontime = ((mobjs(targ).y - mobjs(mo).y) \ mobjs(newmobj).momy) \ t
    End If

    S_StartSound 0, sfx_bospit
End Sub

' travelling cube sound
Private Sub A_SpawnFly(ByVal mo As Long)
    Dim newmobj As Long, fog As Long, targ As Long, r As Long, mtype As Long

    mobjs(mo).reactiontime = mobjs(mo).reactiontime - 1
    If mobjs(mo).reactiontime <> 0 Then Exit Sub ' still flying

    targ = mobjs(mo).target

    ' First spawn teleport fog.
    fog = P_SpawnMobj(mobjs(targ).x, mobjs(targ).y, mobjs(targ).z, MT_SPAWNFIRE)
    S_StartSound fog, sfx_telept

    ' Randomly select monster to spawn.
    r = P_Random()

    ' Probability distribution (kind of :), decreasing likelihood.
    If r < 50 Then
        mtype = MT_TROOP
    ElseIf r < 90 Then
        mtype = MT_SERGEANT
    ElseIf r < 120 Then
        mtype = MT_SHADOWS
    ElseIf r < 130 Then
        mtype = MT_PAIN
    ElseIf r < 160 Then
        mtype = MT_HEAD
    ElseIf r < 162 Then
        mtype = MT_VILE
    ElseIf r < 172 Then
        mtype = MT_UNDEAD
    ElseIf r < 192 Then
        mtype = MT_BABY
    ElseIf r < 222 Then
        mtype = MT_FATSO
    ElseIf r < 246 Then
        mtype = MT_KNIGHT
    Else
        mtype = MT_BRUISER
    End If

    newmobj = P_SpawnMobj(mobjs(targ).x, mobjs(targ).y, mobjs(targ).z, mtype)
    If P_LookForPlayers(newmobj, True) Then P_SetMobjState newmobj, mobjinfo(mobjs(newmobj).info).seestate

    ' telefrag anything in this spot
    P_TeleportMove newmobj, mobjs(newmobj).x, mobjs(newmobj).y

    ' remove self (i.e., cube).
    P_RemoveMobj mo
End Sub

Private Sub A_PlayerScream(ByVal mo As Long)
    Dim sound As Long

    ' Default death sound.
    sound = sfx_pldeth

    If gamemode = commercial And mobjs(mo).health < -50 Then
        ' IF THE PLAYER DIES LESS THAN -50% WITHOUT GIBBING
        sound = sfx_pdiehi
    End If

    S_StartSound mo, sound
End Sub
