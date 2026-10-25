Attribute VB_Name = "p_mobj"
' p_mobj.bas - moving object handling.  Spawn functions (p_mobj.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const STOPSPEED As Long = &H1000&
Private Const FRICTION As Long = &HE800&

' (P_Random() - P_Random()) with the left call made first, as the original
' executable did.
Private Function P_SubRandom() As Long
    Dim r As Long
    r = P_Random()
    P_SubRandom = r - P_Random()
End Function

' P_SetMobjState - returns true if the mobj is still present.
Public Function P_SetMobjState(ByVal mo As Long, ByVal state As Long) As Boolean
    Do
        If state = S_NULL Then
            mobjs(mo).state = S_NULL
            P_RemoveMobj mo
            P_SetMobjState = False
            Exit Function
        End If

        If mo = watchslot Then I_Printf "watch: tic " & gametic & " state " & state & " (func " & thinkers(mo).func & ") by thinker " & currentThinker
        mobjs(mo).state = state
        mobjs(mo).tics = states(state).tics
        mobjs(mo).sprite = states(state).sprite
        mobjs(mo).frame = states(state).frame

        ' Modified handling.
        ' Call action functions when the state is set
        If states(state).action <> AF_NULL Then P_CallMobjAction states(state).action, mo

        state = states(state).nextstate
    Loop While mobjs(mo).tics = 0

    P_SetMobjState = True
End Function

' P_ExplodeMissile
Public Sub P_ExplodeMissile(ByVal mo As Long)
    mobjs(mo).momx = 0
    mobjs(mo).momy = 0
    mobjs(mo).momz = 0

    P_SetMobjState mo, mobjinfo(mobjs(mo).mtype).deathstate

    mobjs(mo).tics = mobjs(mo).tics - (P_Random() And 3)

    If mobjs(mo).tics < 1 Then mobjs(mo).tics = 1

    mobjs(mo).flags = mobjs(mo).flags And (Not MF_MISSILE)

    If mobjinfo(mobjs(mo).info).deathsound <> 0 Then S_StartSound mo, mobjinfo(mobjs(mo).info).deathsound
End Sub

' P_XYMovement
Private Sub P_XYMovement(ByVal mo As Long)
    Dim ptryx As Long, ptryy As Long, player As Long, xmove As Long, ymove As Long
    Dim cl As Long, st As Long, notmoving As Boolean

    If mobjs(mo).momx = 0 And mobjs(mo).momy = 0 Then
        If (mobjs(mo).flags And MF_SKULLFLY) <> 0 Then
            ' the skull slammed into something
            mobjs(mo).flags = mobjs(mo).flags And (Not MF_SKULLFLY)
            mobjs(mo).momx = 0
            mobjs(mo).momy = 0
            mobjs(mo).momz = 0

            P_SetMobjState mo, mobjinfo(mobjs(mo).info).spawnstate
        End If
        Exit Sub
    End If

    player = mobjs(mo).player

    If mobjs(mo).momx > MAXMOVE Then
        mobjs(mo).momx = MAXMOVE
    ElseIf mobjs(mo).momx < -MAXMOVE Then
        mobjs(mo).momx = -MAXMOVE
    End If

    If mobjs(mo).momy > MAXMOVE Then
        mobjs(mo).momy = MAXMOVE
    ElseIf mobjs(mo).momy < -MAXMOVE Then
        mobjs(mo).momy = -MAXMOVE
    End If

    xmove = mobjs(mo).momx
    ymove = mobjs(mo).momy

    Do
        If xmove > MAXMOVE \ 2 Or ymove > MAXMOVE \ 2 Then
            ptryx = mobjs(mo).x + xmove \ 2
            ptryy = mobjs(mo).y + ymove \ 2
            xmove = SAR(xmove, 1)
            ymove = SAR(ymove, 1)
        Else
            ptryx = mobjs(mo).x + xmove
            ptryy = mobjs(mo).y + ymove
            xmove = 0
            ymove = 0
        End If

        If Not P_TryMove(mo, ptryx, ptryy) Then
            ' blocked move
            If mobjs(mo).player >= 0 Then
                ' try to slide along it
                P_SlideMove mo
            ElseIf (mobjs(mo).flags And MF_MISSILE) <> 0 Then
                ' explode a missile
                cl = ceilingline
                If cl <> -1 Then
                    If lines(cl).backsector <> -1 Then
                        If sectors(lines(cl).backsector).ceilingpic = skyflatnum Then
                            ' Hack to prevent missiles exploding against the sky.
                            ' Does not handle sky floors.
                            P_RemoveMobj mo
                            Exit Sub
                        End If
                    End If
                End If
                P_ExplodeMissile mo
            Else
                mobjs(mo).momx = 0
                mobjs(mo).momy = 0
            End If
        End If
    Loop While xmove <> 0 Or ymove <> 0

    ' slow down
    If player >= 0 Then
        If (players(player).cheats And CF_NOMOMENTUM) <> 0 Then
            ' debug option for no sliding at all
            mobjs(mo).momx = 0
            mobjs(mo).momy = 0
            Exit Sub
        End If
    End If

    If (mobjs(mo).flags And (MF_MISSILE Or MF_SKULLFLY)) <> 0 Then Exit Sub ' no friction for missiles ever

    If mobjs(mo).z > mobjs(mo).floorz Then Exit Sub ' no friction when airborne

    If (mobjs(mo).flags And MF_CORPSE) <> 0 Then
        ' do not stop sliding if halfway off a step with some momentum
        If mobjs(mo).momx > FRACUNIT \ 4 Or mobjs(mo).momx < -FRACUNIT \ 4 Or _
           mobjs(mo).momy > FRACUNIT \ 4 Or mobjs(mo).momy < -FRACUNIT \ 4 Then
            If mobjs(mo).floorz <> sectors(subsectors(mobjs(mo).subsector).sector).floorheight Then Exit Sub
        End If
    End If

    ' (C short-circuits "!player || (no forward/side move)")
    If player < 0 Then
        notmoving = True
    Else
        notmoving = PlayerNotMoving(player)
    End If

    If mobjs(mo).momx > -STOPSPEED And mobjs(mo).momx < STOPSPEED And _
       mobjs(mo).momy > -STOPSPEED And mobjs(mo).momy < STOPSPEED And notmoving Then
        ' if in a walking frame, stop moving
        If player >= 0 Then
            st = mobjs(players(player).mo).state
            If st >= S_PLAY_RUN1 And st < S_PLAY_RUN1 + 4 Then P_SetMobjState players(player).mo, S_PLAY
        End If

        mobjs(mo).momx = 0
        mobjs(mo).momy = 0
    Else
        mobjs(mo).momx = FixedMul(mobjs(mo).momx, FRICTION)
        mobjs(mo).momy = FixedMul(mobjs(mo).momy, FRICTION)
    End If
End Sub

Private Function PlayerNotMoving(ByVal player As Long) As Boolean
    PlayerNotMoving = (players(player).cmd.forwardmove = 0 And players(player).cmd.sidemove = 0)
End Function

' P_ZMovement
Private Sub P_ZMovement(ByVal mo As Long)
    Dim dist As Long, delta As Long, p As Long, tgt As Long

    p = mobjs(mo).player

    ' check for smooth step up
    If p >= 0 Then
        If mobjs(mo).z < mobjs(mo).floorz Then
            players(p).viewheight = players(p).viewheight - (mobjs(mo).floorz - mobjs(mo).z)
            players(p).deltaviewheight = SAR(PLAYER_VIEWHEIGHT - players(p).viewheight, 3)
        End If
    End If

    ' adjust height
    mobjs(mo).z = mobjs(mo).z + mobjs(mo).momz

    tgt = mobjs(mo).target
    If (mobjs(mo).flags And MF_FLOAT) <> 0 And tgt <> 0 Then
        ' float down towards target if too close
        If (mobjs(mo).flags And MF_SKULLFLY) = 0 And (mobjs(mo).flags And MF_INFLOAT) = 0 Then
            dist = P_AproxDistance(mobjs(mo).x - mobjs(tgt).x, mobjs(mo).y - mobjs(tgt).y)

            delta = (mobjs(tgt).z + SAR(mobjs(mo).height, 1)) - mobjs(mo).z

            If delta < 0 And dist < -(delta * 3) Then
                mobjs(mo).z = mobjs(mo).z - FLOATSPEED
            ElseIf delta > 0 And dist < (delta * 3) Then
                mobjs(mo).z = mobjs(mo).z + FLOATSPEED
            End If
        End If
    End If

    ' clip movement
    If mobjs(mo).z <= mobjs(mo).floorz Then
        ' hit the floor
        If (mobjs(mo).flags And MF_SKULLFLY) <> 0 Then
            ' the skull slammed into something
            mobjs(mo).momz = -mobjs(mo).momz
        End If

        If mobjs(mo).momz < 0 Then
            If p >= 0 And mobjs(mo).momz < -GRAVITY * 8 Then
                ' Squat down.
                ' Decrease viewheight for a moment after hitting the ground
                ' (hard), and utter appropriate sound.
                players(p).deltaviewheight = SAR(mobjs(mo).momz, 3)
                S_StartSound mo, sfx_oof
            End If
            mobjs(mo).momz = 0
        End If
        mobjs(mo).z = mobjs(mo).floorz

        If (mobjs(mo).flags And MF_MISSILE) <> 0 And (mobjs(mo).flags And MF_NOCLIP) = 0 Then
            P_ExplodeMissile mo
            Exit Sub
        End If
    ElseIf (mobjs(mo).flags And MF_NOGRAVITY) = 0 Then
        If mobjs(mo).momz = 0 Then
            mobjs(mo).momz = -GRAVITY * 2
        Else
            mobjs(mo).momz = mobjs(mo).momz - GRAVITY
        End If
    End If

    If mobjs(mo).z + mobjs(mo).height > mobjs(mo).ceilingz Then
        ' hit the ceiling
        If mobjs(mo).momz > 0 Then mobjs(mo).momz = 0
        mobjs(mo).z = mobjs(mo).ceilingz - mobjs(mo).height

        If (mobjs(mo).flags And MF_SKULLFLY) <> 0 Then
            ' the skull slammed into something
            mobjs(mo).momz = -mobjs(mo).momz
        End If

        If (mobjs(mo).flags And MF_MISSILE) <> 0 And (mobjs(mo).flags And MF_NOCLIP) = 0 Then
            P_ExplodeMissile mo
            Exit Sub
        End If
    End If
End Sub

' P_NightmareRespawn
Private Sub P_NightmareRespawn(ByVal mobj As Long)
    Dim x As Long, y As Long, z As Long, ss As Long, mo As Long
    Dim mthing As mapthing_t

    x = mobjs(mobj).spawnpoint.x * FRACUNIT
    y = mobjs(mobj).spawnpoint.y * FRACUNIT

    ' somthing is occupying it's position?
    If Not P_CheckPosition(mobj, x, y) Then Exit Sub ' no respwan

    ' spawn a teleport fog at old spot because of removal of the body?
    mo = P_SpawnMobj(mobjs(mobj).x, mobjs(mobj).y, _
                     sectors(subsectors(mobjs(mobj).subsector).sector).floorheight, MT_TFOG)
    ' initiate teleport sound
    S_StartSound mo, sfx_telept

    ' spawn a teleport fog at the new spot
    ss = R_PointInSubsector(x, y)

    mo = P_SpawnMobj(x, y, sectors(subsectors(ss).sector).floorheight, MT_TFOG)

    S_StartSound mo, sfx_telept

    ' spawn the new monster (copy, so no ByRef element of mobjs() is held)
    mthing = mobjs(mobj).spawnpoint

    ' spawn it
    If (mobjinfo(mobjs(mobj).info).flags And MF_SPAWNCEILING) <> 0 Then
        z = ONCEILINGZ
    Else
        z = ONFLOORZ
    End If

    ' inherit attributes from deceased one
    mo = P_SpawnMobj(x, y, z, mobjs(mobj).mtype)
    mobjs(mo).spawnpoint = mthing
    mobjs(mo).angle = ANG45 * (mthing.angle \ 45)

    If (mthing.options And MTF_AMBUSH) <> 0 Then mobjs(mo).flags = mobjs(mo).flags Or MF_AMBUSH

    mobjs(mo).reactiontime = 18

    ' remove the old monster,
    P_RemoveMobj mobj
End Sub

' P_MobjThinker
Public Sub P_MobjThinker(ByVal mobj As Long)
    ' momentum movement
    If mobjs(mobj).momx <> 0 Or mobjs(mobj).momy <> 0 Or (mobjs(mobj).flags And MF_SKULLFLY) <> 0 Then
        P_XYMovement mobj

        ' FIXME: decent NOP/NULL/Nil function pointer please.
        If thinkers(mobj).func = TF_REMOVED Then Exit Sub ' mobj was removed
    End If
    If mobjs(mobj).z <> mobjs(mobj).floorz Or mobjs(mobj).momz <> 0 Then
        P_ZMovement mobj

        ' FIXME: decent NOP/NULL/Nil function pointer please.
        If thinkers(mobj).func = TF_REMOVED Then Exit Sub ' mobj was removed
    End If

    ' cycle through states, calling action functions at transitions
    If mobjs(mobj).tics <> -1 Then
        mobjs(mobj).tics = mobjs(mobj).tics - 1

        ' you can cycle through multiple states in a tic
        If mobjs(mobj).tics = 0 Then
            If Not P_SetMobjState(mobj, states(mobjs(mobj).state).nextstate) Then Exit Sub ' freed itself
        End If
    Else
        ' check for nightmare respawn
        If (mobjs(mobj).flags And MF_COUNTKILL) = 0 Then Exit Sub

        If Not respawnmonsters Then Exit Sub

        mobjs(mobj).movecount = mobjs(mobj).movecount + 1

        If mobjs(mobj).movecount < 12 * 35 Then Exit Sub

        If (leveltime And 31) <> 0 Then Exit Sub

        If P_Random() > 4 Then Exit Sub

        P_NightmareRespawn mobj
    End If
End Sub

' P_SpawnMobj
Public Function P_SpawnMobj(ByVal x As Long, ByVal y As Long, ByVal z As Long, ByVal mtype As Long) As Long
    Dim mobj As Long, st As Long, ss As Long

    mobj = P_NewThinker()

    mobjs(mobj).mtype = mtype
    mobjs(mobj).info = mtype
    mobjs(mobj).x = x
    mobjs(mobj).y = y
    mobjs(mobj).radius = mobjinfo(mtype).radius
    mobjs(mobj).height = mobjinfo(mtype).height
    mobjs(mobj).flags = mobjinfo(mtype).flags
    mobjs(mobj).health = mobjinfo(mtype).spawnhealth
    mobjs(mobj).player = -1

    If gameskill <> sk_nightmare Then mobjs(mobj).reactiontime = mobjinfo(mtype).reactiontime

    mobjs(mobj).lastlook = P_Random() Mod MAXPLAYERS
    ' do not set the state with P_SetMobjState, because action routines can
    ' not be called yet
    st = mobjinfo(mtype).spawnstate

    mobjs(mobj).state = st
    mobjs(mobj).tics = states(st).tics
    mobjs(mobj).sprite = states(st).sprite
    mobjs(mobj).frame = states(st).frame

    ' set subsector and/or block links
    P_SetThingPosition mobj

    ss = subsectors(mobjs(mobj).subsector).sector
    mobjs(mobj).floorz = sectors(ss).floorheight
    mobjs(mobj).ceilingz = sectors(ss).ceilingheight

    If z = ONFLOORZ Then
        mobjs(mobj).z = mobjs(mobj).floorz
    ElseIf z = ONCEILINGZ Then
        mobjs(mobj).z = mobjs(mobj).ceilingz - mobjinfo(mtype).height
    Else
        mobjs(mobj).z = z
    End If

    thinkers(mobj).func = TF_MOBJ

    P_AddThinker mobj

    If mobj = watchslot Then I_Printf "watch: tic " & gametic & " spawn type " & mtype & " at " & (x \ 65536) & "," & (y \ 65536) & " sector " & ss & " by thinker " & currentThinker & " (" & dbgPhase & ")"

    P_SpawnMobj = mobj
End Function

' P_RemoveMobj
Public Sub P_RemoveMobj(ByVal mobj As Long)
    If mobj = watchslot Then I_Printf "watch: tic " & gametic & " remove (func " & thinkers(mobj).func & ") by thinker " & currentThinker & " state " & mobjs(mobj).state
    If checklinks Then
        dbgUnsetBy = "P_RemoveMobj(" & mobj & ") thinker " & currentThinker & " type " & mobjs(currentThinker).mtype & " state " & mobjs(currentThinker).state
        If thinkers(mobj).func <> TF_MOBJ Then Err.Raise vbObjectError + 2, , "P_RemoveMobj on dead slot " & mobj & " (type " & mobjs(mobj).mtype & ") by " & dbgUnsetBy & " at gametic " & gametic
    End If
    If (mobjs(mobj).flags And MF_SPECIAL) <> 0 And (mobjs(mobj).flags And MF_DROPPED) = 0 And _
       mobjs(mobj).mtype <> MT_INV And mobjs(mobj).mtype <> MT_INS Then
        itemrespawnque(iquehead) = mobjs(mobj).spawnpoint
        itemrespawntime(iquehead) = leveltime
        iquehead = (iquehead + 1) And (ITEMQUESIZE - 1)

        ' lose one off the end?
        If iquehead = iquetail Then iquetail = (iquetail + 1) And (ITEMQUESIZE - 1)
    End If

    ' unlink from sector and block lists
    P_UnsetThingPosition mobj

    ' stop any playing sound
    S_StopSound mobj

    ' free block
    P_RemoveThinker mobj
End Sub

' P_RespawnSpecials
Public Sub P_RespawnSpecials()
    Dim x As Long, y As Long, z As Long, ss As Long, mo As Long, i As Long
    Dim mthing As mapthing_t

    ' only respawn items in deathmatch
    If deathmatch <> 2 Then Exit Sub

    ' nothing left to respawn?
    If iquehead = iquetail Then Exit Sub

    ' wait at least 30 seconds
    If leveltime - itemrespawntime(iquetail) < 30 * 35 Then Exit Sub

    mthing = itemrespawnque(iquetail)

    x = mthing.x * FRACUNIT
    y = mthing.y * FRACUNIT

    ' spawn a teleport fog at the new spot
    ss = R_PointInSubsector(x, y)
    mo = P_SpawnMobj(x, y, sectors(subsectors(ss).sector).floorheight, MT_IFOG)
    S_StartSound mo, sfx_itmbk

    ' find which type to spawn
    For i = 0 To NUMMOBJTYPES - 1
        If mthing.typ = mobjinfo(i).doomednum Then Exit For
    Next

    ' spawn it
    If (mobjinfo(i).flags And MF_SPAWNCEILING) <> 0 Then
        z = ONCEILINGZ
    Else
        z = ONFLOORZ
    End If

    mo = P_SpawnMobj(x, y, z, i)
    mobjs(mo).spawnpoint = mthing
    mobjs(mo).angle = ANG45 * (mthing.angle \ 45)

    ' pull it from the que
    iquetail = (iquetail + 1) And (ITEMQUESIZE - 1)
End Sub

' P_SpawnPlayer - called when a player is spawned on the level.  Most of the
' player structure stays unchanged between levels.
Public Sub P_SpawnPlayer(mthing As mapthing_t)
    Dim p As Long, x As Long, y As Long, z As Long, mobj As Long, i As Long

    ' not playing?
    If Not playeringame(mthing.typ - 1) Then Exit Sub

    p = mthing.typ - 1

    If players(p).playerstate = PST_REBORN Then G_PlayerReborn mthing.typ - 1

    x = mthing.x * FRACUNIT
    y = mthing.y * FRACUNIT
    z = ONFLOORZ
    mobj = P_SpawnMobj(x, y, z, MT_PLAYER)

    ' set color translations for player sprites
    If mthing.typ > 1 Then mobjs(mobj).flags = mobjs(mobj).flags Or ((mthing.typ - 1) * pow2(MF_TRANSSHIFT))

    mobjs(mobj).angle = ANG45 * (mthing.angle \ 45)
    mobjs(mobj).player = p
    mobjs(mobj).health = players(p).health

    players(p).mo = mobj
    players(p).playerstate = PST_LIVE
    players(p).refire = 0
    players(p).message = ""
    players(p).damagecount = 0
    players(p).bonuscount = 0
    players(p).extralight = 0
    players(p).fixedcolormap = 0
    players(p).viewheight = PLAYER_VIEWHEIGHT

    ' setup gun psprite
    P_SetupPsprites p

    ' give all cards in death match mode
    If deathmatch <> 0 Then
        For i = 0 To NUMCARDS - 1
            players(p).cards(i) = True
        Next
    End If

    If mthing.typ - 1 = consoleplayer Then
        ' wake up the status bar
        ST_Start
        ' wake up the heads up text
        HU_Start
    End If
End Sub

' P_SpawnMapThing - the fields of the mapthing should already be in host
' byte order.
Public Sub P_SpawnMapThing(mthing As mapthing_t)
    Dim i As Long, bit As Long, mobj As Long, x As Long, y As Long, z As Long

    ' count deathmatch start positions
    If mthing.typ = 11 Then
        If deathmatch_p < MAX_DM_STARTS Then
            deathmatchstarts(deathmatch_p) = mthing
            deathmatch_p = deathmatch_p + 1
        End If
        Exit Sub
    End If

    ' check for players specially
    If mthing.typ <= 4 Then
        If mthing.typ < 1 Then Exit Sub
        ' save spots for respawning in network games
        playerstarts(mthing.typ - 1) = mthing
        If deathmatch = 0 Then P_SpawnPlayer mthing
        Exit Sub
    End If

    ' check for apropriate skill level
    If Not netgame And (mthing.options And 16) <> 0 Then Exit Sub

    If gameskill = sk_baby Then
        bit = 1
    ElseIf gameskill = sk_nightmare Then
        bit = 4
    Else
        bit = pow2(gameskill - 1)
    End If

    If (mthing.options And bit) = 0 Then Exit Sub

    ' find which type to spawn
    For i = 0 To NUMMOBJTYPES - 1
        If mthing.typ = mobjinfo(i).doomednum Then Exit For
    Next

    If i = NUMMOBJTYPES Then
        I_Error "P_SpawnMapThing: Unknown type " & mthing.typ & " at (" & mthing.x & ", " & mthing.y & ")"
    End If

    ' don't spawn keycards and players in deathmatch
    If deathmatch <> 0 And (mobjinfo(i).flags And MF_NOTDMATCH) <> 0 Then Exit Sub

    ' don't spawn any monsters if -nomonsters
    If nomonsters And (i = MT_SKULL Or (mobjinfo(i).flags And MF_COUNTKILL) <> 0) Then Exit Sub

    ' spawn it
    x = mthing.x * FRACUNIT
    y = mthing.y * FRACUNIT

    If (mobjinfo(i).flags And MF_SPAWNCEILING) <> 0 Then
        z = ONCEILINGZ
    Else
        z = ONFLOORZ
    End If

    mobj = P_SpawnMobj(x, y, z, i)
    mobjs(mobj).spawnpoint = mthing

    If mobjs(mobj).tics > 0 Then mobjs(mobj).tics = 1 + (P_Random() Mod mobjs(mobj).tics)
    If (mobjs(mobj).flags And MF_COUNTKILL) <> 0 Then totalkills = totalkills + 1
    If (mobjs(mobj).flags And MF_COUNTITEM) <> 0 Then totalitems = totalitems + 1

    mobjs(mobj).angle = ANG45 * (mthing.angle \ 45)
    If (mthing.options And MTF_AMBUSH) <> 0 Then mobjs(mobj).flags = mobjs(mobj).flags Or MF_AMBUSH
End Sub

' GAME SPAWN FUNCTIONS

' P_SpawnPuff
Public Sub P_SpawnPuff(ByVal x As Long, ByVal y As Long, ByVal z As Long)
    Dim th As Long

    z = z + (P_SubRandom() * 1024)

    th = P_SpawnMobj(x, y, z, MT_PUFF)
    mobjs(th).momz = FRACUNIT
    mobjs(th).tics = mobjs(th).tics - (P_Random() And 3)

    If mobjs(th).tics < 1 Then mobjs(th).tics = 1

    ' don't make punches spark on the wall
    If attackrange = MELEERANGE Then P_SetMobjState th, S_PUFF3
End Sub

' P_SpawnBlood
Public Sub P_SpawnBlood(ByVal x As Long, ByVal y As Long, ByVal z As Long, ByVal damage As Long)
    Dim th As Long

    z = z + (P_SubRandom() * 1024)
    th = P_SpawnMobj(x, y, z, MT_BLOOD)
    mobjs(th).momz = FRACUNIT * 2
    mobjs(th).tics = mobjs(th).tics - (P_Random() And 3)

    If mobjs(th).tics < 1 Then mobjs(th).tics = 1

    If damage <= 12 And damage >= 9 Then
        P_SetMobjState th, S_BLOOD2
    ElseIf damage < 9 Then
        P_SetMobjState th, S_BLOOD3
    End If
End Sub

' P_CheckMissileSpawn - moves the missile forward a bit and possibly
' explodes it right there.
Public Sub P_CheckMissileSpawn(ByVal th As Long)
    mobjs(th).tics = mobjs(th).tics - (P_Random() And 3)
    If mobjs(th).tics < 1 Then mobjs(th).tics = 1

    ' move a little forward so an angle can be computed if it immediately
    ' explodes
    mobjs(th).x = mobjs(th).x + SAR(mobjs(th).momx, 1)
    mobjs(th).y = mobjs(th).y + SAR(mobjs(th).momy, 1)
    mobjs(th).z = mobjs(th).z + SAR(mobjs(th).momz, 1)

    If Not P_TryMove(th, mobjs(th).x, mobjs(th).y) Then P_ExplodeMissile th
End Sub

' P_SpawnMissile
Public Function P_SpawnMissile(ByVal source As Long, ByVal dest As Long, ByVal mtype As Long) As Long
    Dim th As Long, an As Long, dist As Long, speed As Long

    th = P_SpawnMobj(mobjs(source).x, mobjs(source).y, mobjs(source).z + 4 * 8 * FRACUNIT, mtype)

    If mobjinfo(mtype).seesound <> 0 Then S_StartSound th, mobjinfo(mtype).seesound

    mobjs(th).target = source                    ' where it came from
    an = R_PointToAngle2(mobjs(source).x, mobjs(source).y, mobjs(dest).x, mobjs(dest).y)

    ' fuzzy player
    If (mobjs(dest).flags And MF_SHADOW) <> 0 Then an = an + (P_SubRandom() * &H100000)

    mobjs(th).angle = an
    an = FINEANGLE(an)
    speed = mobjinfo(mtype).speed
    mobjs(th).momx = FixedMul(speed, finecosine(an))
    mobjs(th).momy = FixedMul(speed, finesine(an))

    dist = P_AproxDistance(mobjs(dest).x - mobjs(source).x, mobjs(dest).y - mobjs(source).y)
    dist = dist \ speed

    If dist < 1 Then dist = 1

    mobjs(th).momz = (mobjs(dest).z - mobjs(source).z) \ dist
    P_CheckMissileSpawn th

    P_SpawnMissile = th
End Function

' P_SpawnPlayerMissile - tries to aim at a nearby monster
Public Sub P_SpawnPlayerMissile(ByVal source As Long, ByVal mtype As Long)
    Dim th As Long, an As Long, x As Long, y As Long, z As Long, slope As Long, speed As Long

    ' see which target is to be aimed at
    an = mobjs(source).angle
    slope = P_AimLineAttack(source, an, 16 * 64 * FRACUNIT)

    If linetarget = 0 Then
        an = an + &H4000000
        slope = P_AimLineAttack(source, an, 16 * 64 * FRACUNIT)

        If linetarget = 0 Then
            an = an - &H8000000
            slope = P_AimLineAttack(source, an, 16 * 64 * FRACUNIT)
        End If

        If linetarget = 0 Then
            an = mobjs(source).angle
            slope = 0
        End If
    End If

    x = mobjs(source).x
    y = mobjs(source).y
    z = mobjs(source).z + 4 * 8 * FRACUNIT

    th = P_SpawnMobj(x, y, z, mtype)

    If mobjinfo(mtype).seesound <> 0 Then S_StartSound th, mobjinfo(mtype).seesound

    mobjs(th).target = source
    mobjs(th).angle = an
    speed = mobjinfo(mtype).speed
    mobjs(th).momx = FixedMul(speed, finecosine(FINEANGLE(an)))
    mobjs(th).momy = FixedMul(speed, finesine(FINEANGLE(an)))
    mobjs(th).momz = FixedMul(speed, slope)

    P_CheckMissileSpawn th
End Sub
