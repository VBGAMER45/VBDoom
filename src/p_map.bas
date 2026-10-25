Attribute VB_Name = "p_map"
' p_map.bas - movement, collision handling.  Shooting and aiming (p_map.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private tmbbox(0 To 3) As Long
Private tmthing As Long
Private tmflags As Long
Private tmx As Long
Private tmy As Long

' If "floatok" true, move would be ok if within "tmfloorz - tmceilingz".
' (floatok, tmfloorz, tmceilingz, ceilingline are globals in doomdef)
Private tmdropoffz As Long

' keep track of special lines as they are hit, but don't process them until
' the move is proven valid (C: MAXSPECIALCROSS = 8, overflowing into memory)
Public spechit() As Long                         ' (read by P_Move in p_enemy)
Public numspechit As Long
Private spechitInit As Boolean

' slide
Private bestslidefrac As Long
Private secondslidefrac As Long
Private bestslideline As Long
Private secondslideline As Long
Private slidemo As Long
Private tmxmove As Long
Private tmymove As Long

' aim / shoot (linetarget, attackrange, topslope, bottomslope in doomdef)
Private shootthing As Long
Private shootz As Long                           ' height if not aiming up or down
Private la_damage As Long
Private aimslope As Long

Private usething As Long

Private bombsource As Long
Private bombspot As Long
Private bombdamage As Long

Private crushchange As Boolean
Private nofit As Boolean

Private Function P_SubRandom() As Long
    Dim r As Long
    r = P_Random()
    P_SubRandom = r - P_Random()
End Function

Private Sub AddSpecHit(ByVal ld As Long)
    If Not spechitInit Then
        ReDim spechit(0 To 15)
        spechitInit = True
    End If
    If numspechit > UBound(spechit) Then ReDim Preserve spechit(0 To numspechit * 2)
    spechit(numspechit) = ld
    numspechit = numspechit + 1
End Sub

' TELEPORT MOVE

' PIT_StompThing
Public Function PIT_StompThing(ByVal thing As Long) As Boolean
    Dim blockdist As Long

    If (mobjs(thing).flags And MF_SHOOTABLE) = 0 Then
        PIT_StompThing = True
        Exit Function
    End If

    blockdist = mobjs(thing).radius + mobjs(tmthing).radius

    If Abs(mobjs(thing).x - tmx) >= blockdist Or Abs(mobjs(thing).y - tmy) >= blockdist Then
        ' didn't hit it
        PIT_StompThing = True
        Exit Function
    End If

    ' don't clip against self
    If thing = tmthing Then
        PIT_StompThing = True
        Exit Function
    End If

    ' monsters don't stomp things except on boss level
    If mobjs(tmthing).player < 0 And gamemap <> 30 Then
        PIT_StompThing = False
        Exit Function
    End If

    P_DamageMobj thing, tmthing, tmthing, 10000

    PIT_StompThing = True
End Function

' P_TeleportMove
Public Function P_TeleportMove(ByVal thing As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim xl As Long, xh As Long, yl As Long, yh As Long, bx As Long, by As Long
    Dim newsubsec As Long
    If checklinks Then
        dbgUnsetBy = "P_TeleportMove(" & thing & ") thinker " & currentThinker & " type " & mobjs(currentThinker).mtype & " state " & mobjs(currentThinker).state
        If thinkers(thing).func <> TF_MOBJ Then Err.Raise vbObjectError + 2, , "P_TeleportMove on dead slot " & thing & " (type " & mobjs(thing).mtype & ") by " & dbgUnsetBy & " at gametic " & gametic
    End If

    ' kill anything occupying the position
    tmthing = thing
    tmflags = mobjs(thing).flags

    tmx = x
    tmy = y

    tmbbox(BOXTOP) = y + mobjs(tmthing).radius
    tmbbox(BOXBOTTOM) = y - mobjs(tmthing).radius
    tmbbox(BOXRIGHT) = x + mobjs(tmthing).radius
    tmbbox(BOXLEFT) = x - mobjs(tmthing).radius

    newsubsec = R_PointInSubsector(x, y)
    ceilingline = -1

    ' The base floor/ceiling is from the subsector that contains the point.
    ' Any contacted lines the step closer together will adjust them.
    tmfloorz = sectors(subsectors(newsubsec).sector).floorheight
    tmdropoffz = tmfloorz
    tmceilingz = sectors(subsectors(newsubsec).sector).ceilingheight

    validcount = validcount + 1
    numspechit = 0

    ' stomp on any things contacted
    xl = SAR(tmbbox(BOXLEFT) - bmaporgx - MAXRADIUS, MAPBLOCKSHIFT)
    xh = SAR(tmbbox(BOXRIGHT) - bmaporgx + MAXRADIUS, MAPBLOCKSHIFT)
    yl = SAR(tmbbox(BOXBOTTOM) - bmaporgy - MAXRADIUS, MAPBLOCKSHIFT)
    yh = SAR(tmbbox(BOXTOP) - bmaporgy + MAXRADIUS, MAPBLOCKSHIFT)

    For bx = xl To xh
        For by = yl To yh
            If Not P_BlockThingsIterator(bx, by, ITER_STOMPTHING) Then
                P_TeleportMove = False
                Exit Function
            End If
        Next
    Next

    ' the move is ok, so link the thing into its new position
    P_UnsetThingPosition thing

    mobjs(thing).floorz = tmfloorz
    mobjs(thing).ceilingz = tmceilingz
    mobjs(thing).x = x
    mobjs(thing).y = y

    P_SetThingPosition thing

    P_TeleportMove = True
End Function

' MOVEMENT ITERATOR FUNCTIONS

' PIT_CheckLine - adjusts tmfloorz and tmceilingz as lines are contacted
Public Function PIT_CheckLine(ByVal ld As Long) As Boolean
    If tmbbox(BOXRIGHT) <= lines(ld).bbox(BOXLEFT) Or tmbbox(BOXLEFT) >= lines(ld).bbox(BOXRIGHT) Or _
       tmbbox(BOXTOP) <= lines(ld).bbox(BOXBOTTOM) Or tmbbox(BOXBOTTOM) >= lines(ld).bbox(BOXTOP) Then
        PIT_CheckLine = True
        Exit Function
    End If

    If P_BoxOnLineSide(tmbbox, ld) <> -1 Then
        PIT_CheckLine = True
        Exit Function
    End If

    ' A line has been hit

    ' The moving thing's destination position will cross the given line.
    ' If this should not be allowed, return false.
    ' If the line is special, keep track of it to process later if the move
    '  is proven ok.
    ' NOTE: specials are NOT sorted by order, so two special lines that are
    '  only 8 pixels apart could be crossed in either order.

    If lines(ld).backsector = -1 Then
        PIT_CheckLine = False                    ' one sided line
        Exit Function
    End If

    If (mobjs(tmthing).flags And MF_MISSILE) = 0 Then
        If (lines(ld).flags And ML_BLOCKING) <> 0 Then
            PIT_CheckLine = False                ' explicitly blocking everything
            Exit Function
        End If

        If mobjs(tmthing).player < 0 And (lines(ld).flags And ML_BLOCKMONSTERS) <> 0 Then
            PIT_CheckLine = False                ' block monsters only
            Exit Function
        End If
    End If

    ' set openrange, opentop, openbottom
    P_LineOpening ld

    ' adjust floor / ceiling heights
    If opentop < tmceilingz Then
        tmceilingz = opentop
        ceilingline = ld
    End If

    If openbottom > tmfloorz Then tmfloorz = openbottom

    If lowfloor < tmdropoffz Then tmdropoffz = lowfloor

    ' if contacted a special line, add it to the list
    If lines(ld).special <> 0 Then AddSpecHit ld

    PIT_CheckLine = True
End Function

' PIT_CheckThing
Public Function PIT_CheckThing(ByVal thing As Long) As Boolean
    Dim blockdist As Long, solid As Boolean, damage As Long, tgt As Long

    If (mobjs(thing).flags And (MF_SOLID Or MF_SPECIAL Or MF_SHOOTABLE)) = 0 Then
        PIT_CheckThing = True
        Exit Function
    End If

    blockdist = mobjs(thing).radius + mobjs(tmthing).radius

    If Abs(mobjs(thing).x - tmx) >= blockdist Or Abs(mobjs(thing).y - tmy) >= blockdist Then
        ' didn't hit it
        PIT_CheckThing = True
        Exit Function
    End If

    ' don't clip against self
    If thing = tmthing Then
        PIT_CheckThing = True
        Exit Function
    End If

    ' check for skulls slamming into things
    If (mobjs(tmthing).flags And MF_SKULLFLY) <> 0 Then
        damage = ((P_Random() Mod 8) + 1) * mobjinfo(mobjs(tmthing).info).damage

        P_DamageMobj thing, tmthing, tmthing, damage

        mobjs(tmthing).flags = mobjs(tmthing).flags And (Not MF_SKULLFLY)
        mobjs(tmthing).momx = 0
        mobjs(tmthing).momy = 0
        mobjs(tmthing).momz = 0

        P_SetMobjState tmthing, mobjinfo(mobjs(tmthing).info).spawnstate

        PIT_CheckThing = False                   ' stop moving
        Exit Function
    End If

    ' missiles can hit other things
    If (mobjs(tmthing).flags And MF_MISSILE) <> 0 Then
        ' see if it went over / under
        If mobjs(tmthing).z > mobjs(thing).z + mobjs(thing).height Then
            PIT_CheckThing = True                ' overhead
            Exit Function
        End If
        If mobjs(tmthing).z + mobjs(tmthing).height < mobjs(thing).z Then
            PIT_CheckThing = True                ' underneath
            Exit Function
        End If

        tgt = mobjs(tmthing).target
        If tgt <> 0 Then
            If mobjs(tgt).mtype = mobjs(thing).mtype Or _
               (mobjs(tgt).mtype = MT_KNIGHT And mobjs(thing).mtype = MT_BRUISER) Or _
               (mobjs(tgt).mtype = MT_BRUISER And mobjs(thing).mtype = MT_KNIGHT) Then
                ' Don't hit same species as originator.
                If thing = tgt Then
                    PIT_CheckThing = True
                    Exit Function
                End If

                If mobjs(thing).mtype <> MT_PLAYER Then
                    ' Explode, but do no damage.
                    ' Let players missile other players.
                    PIT_CheckThing = False
                    Exit Function
                End If
            End If
        End If

        If (mobjs(thing).flags And MF_SHOOTABLE) = 0 Then
            ' didn't do any damage
            PIT_CheckThing = ((mobjs(thing).flags And MF_SOLID) = 0)
            Exit Function
        End If

        ' damage / explode
        damage = ((P_Random() Mod 8) + 1) * mobjinfo(mobjs(tmthing).info).damage
        P_DamageMobj thing, tmthing, mobjs(tmthing).target, damage

        ' don't traverse any more
        PIT_CheckThing = False
        Exit Function
    End If

    ' check for special pickup
    If (mobjs(thing).flags And MF_SPECIAL) <> 0 Then
        solid = ((mobjs(thing).flags And MF_SOLID) <> 0)
        If (tmflags And MF_PICKUP) <> 0 Then
            ' can remove thing
            P_TouchSpecialThing thing, tmthing
        End If
        PIT_CheckThing = Not solid
        Exit Function
    End If

    PIT_CheckThing = ((mobjs(thing).flags And MF_SOLID) = 0)
End Function

' MOVEMENT CLIPPING

' P_CheckPosition - this is purely informative, nothing is modified (except
' things picked up).
'
' in:
'  a mobj_t (can be valid or invalid)
'  a position to be checked
'   (doesn't need to be related to the mobj_t->x,y)
'
' during:
'  special things are touched if MF_PICKUP
'  early out on solid lines?
'
' out:
'  newsubsec
'  floorz
'  ceilingz
'  tmdropoffz
'   the lowest point contacted
'   (monsters won't move to a dropoff)
'  speciallines[]
'  numspeciallines
Public Function P_CheckPosition(ByVal thing As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim xl As Long, xh As Long, yl As Long, yh As Long, bx As Long, by As Long
    Dim newsubsec As Long

    tmthing = thing
    tmflags = mobjs(thing).flags

    tmx = x
    tmy = y

    tmbbox(BOXTOP) = y + mobjs(tmthing).radius
    tmbbox(BOXBOTTOM) = y - mobjs(tmthing).radius
    tmbbox(BOXRIGHT) = x + mobjs(tmthing).radius
    tmbbox(BOXLEFT) = x - mobjs(tmthing).radius

    newsubsec = R_PointInSubsector(x, y)
    ceilingline = -1

    ' The base floor / ceiling is from the subsector that contains the point.
    ' Any contacted lines the step closer together will adjust them.
    tmfloorz = sectors(subsectors(newsubsec).sector).floorheight
    tmdropoffz = tmfloorz
    tmceilingz = sectors(subsectors(newsubsec).sector).ceilingheight

    validcount = validcount + 1
    numspechit = 0

    If (tmflags And MF_NOCLIP) <> 0 Then
        P_CheckPosition = True
        Exit Function
    End If

    ' Check things first, possibly picking things up.
    ' The bounding box is extended by MAXRADIUS because mobj_ts are grouped
    ' into mapblocks based on their origin point, and can overlap into
    ' adjacent blocks by up to MAXRADIUS units.
    xl = SAR(tmbbox(BOXLEFT) - bmaporgx - MAXRADIUS, MAPBLOCKSHIFT)
    xh = SAR(tmbbox(BOXRIGHT) - bmaporgx + MAXRADIUS, MAPBLOCKSHIFT)
    yl = SAR(tmbbox(BOXBOTTOM) - bmaporgy - MAXRADIUS, MAPBLOCKSHIFT)
    yh = SAR(tmbbox(BOXTOP) - bmaporgy + MAXRADIUS, MAPBLOCKSHIFT)

    For bx = xl To xh
        For by = yl To yh
            If Not P_BlockThingsIterator(bx, by, ITER_CHECKTHING) Then
                P_CheckPosition = False
                Exit Function
            End If
        Next
    Next

    ' check lines
    xl = SAR(tmbbox(BOXLEFT) - bmaporgx, MAPBLOCKSHIFT)
    xh = SAR(tmbbox(BOXRIGHT) - bmaporgx, MAPBLOCKSHIFT)
    yl = SAR(tmbbox(BOXBOTTOM) - bmaporgy, MAPBLOCKSHIFT)
    yh = SAR(tmbbox(BOXTOP) - bmaporgy, MAPBLOCKSHIFT)

    For bx = xl To xh
        For by = yl To yh
            If Not P_BlockLinesIterator(bx, by, ITER_CHECKLINE) Then
                P_CheckPosition = False
                Exit Function
            End If
        Next
    Next

    P_CheckPosition = True
End Function

' P_TryMove - attempt to move to a new position, crossing special lines
' unless MF_TELEPORT is set.
Public Function P_TryMove(ByVal thing As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim oldx As Long, oldy As Long, side As Long, oldside As Long, ld As Long
    If checklinks Then
        dbgUnsetBy = "P_TryMove(" & thing & ") thinker " & currentThinker & " type " & mobjs(currentThinker).mtype & " state " & mobjs(currentThinker).state
        If thinkers(thing).func <> TF_MOBJ Then Err.Raise vbObjectError + 2, , "P_TryMove on dead slot " & thing & " (type " & mobjs(thing).mtype & ") by " & dbgUnsetBy & " at gametic " & gametic
    End If

    floatok = False
    If Not P_CheckPosition(thing, x, y) Then
        P_TryMove = False                        ' solid wall or thing
        Exit Function
    End If

    If (mobjs(thing).flags And MF_NOCLIP) = 0 Then
        If tmceilingz - tmfloorz < mobjs(thing).height Then
            P_TryMove = False                    ' doesn't fit
            Exit Function
        End If

        floatok = True

        If (mobjs(thing).flags And MF_TELEPORT) = 0 And tmceilingz - mobjs(thing).z < mobjs(thing).height Then
            P_TryMove = False                    ' mobj must lower itself to fit
            Exit Function
        End If

        If (mobjs(thing).flags And MF_TELEPORT) = 0 And tmfloorz - mobjs(thing).z > 24 * FRACUNIT Then
            P_TryMove = False                    ' too big a step up
            Exit Function
        End If

        If (mobjs(thing).flags And (MF_DROPOFF Or MF_FLOAT)) = 0 And tmfloorz - tmdropoffz > 24 * FRACUNIT Then
            P_TryMove = False                    ' don't stand over a dropoff
            Exit Function
        End If
    End If

    ' the move is ok, so link the thing into its new position
    P_UnsetThingPosition thing

    oldx = mobjs(thing).x
    oldy = mobjs(thing).y
    mobjs(thing).floorz = tmfloorz
    mobjs(thing).ceilingz = tmceilingz
    mobjs(thing).x = x
    mobjs(thing).y = y

    P_SetThingPosition thing

    ' if any special lines were hit, do the effect
    If (mobjs(thing).flags And (MF_TELEPORT Or MF_NOCLIP)) = 0 Then
        Do While numspechit > 0
            numspechit = numspechit - 1
            ' see if the line was crossed
            ld = spechit(numspechit)
            side = P_PointOnLineSide(mobjs(thing).x, mobjs(thing).y, ld)
            oldside = P_PointOnLineSide(oldx, oldy, ld)
            If side <> oldside Then
                If lines(ld).special <> 0 Then P_CrossSpecialLine ld, oldside, thing
            End If
        Loop
    End If

    P_TryMove = True
End Function

' P_ThingHeightClip - takes a valid thing and adjusts the thing->floorz,
' thing->ceilingz, and possibly thing->z.  This is called for all nearby
' monsters whenever a sector changes height.  If the thing doesn't fit, the
' z will be set to the lowest value and false will be returned.
Private Function P_ThingHeightClip(ByVal thing As Long) As Boolean
    Dim onfloor As Boolean

    onfloor = (mobjs(thing).z = mobjs(thing).floorz)

    P_CheckPosition thing, mobjs(thing).x, mobjs(thing).y
    ' what about stranding a monster partially off an edge?

    mobjs(thing).floorz = tmfloorz
    mobjs(thing).ceilingz = tmceilingz

    If onfloor Then
        ' walking monsters rise and fall with the floor
        mobjs(thing).z = mobjs(thing).floorz
    Else
        ' don't adjust a floating monster unless forced to
        If mobjs(thing).z + mobjs(thing).height > mobjs(thing).ceilingz Then
            mobjs(thing).z = mobjs(thing).ceilingz - mobjs(thing).height
        End If
    End If

    If mobjs(thing).ceilingz - mobjs(thing).floorz < mobjs(thing).height Then
        P_ThingHeightClip = False
    Else
        P_ThingHeightClip = True
    End If
End Function

' SLIDE MOVE
' Allows the player to slide along any angled walls.

' P_HitSlideLine - adjusts the xmove / ymove so that the next move will slide
' along the wall.
Private Sub P_HitSlideLine(ByVal ld As Long)
    Dim side As Long, lineangle As Long, moveangle As Long, deltaangle As Long
    Dim movelen As Long, newlen As Long

    If lines(ld).slopetype = ST_HORIZONTAL Then
        tmymove = 0
        Exit Sub
    End If

    If lines(ld).slopetype = ST_VERTICAL Then
        tmxmove = 0
        Exit Sub
    End If

    side = P_PointOnLineSide(mobjs(slidemo).x, mobjs(slidemo).y, ld)

    lineangle = R_PointToAngle2(0, 0, lines(ld).dx, lines(ld).dy)

    If side = 1 Then lineangle = lineangle + ANG180

    moveangle = R_PointToAngle2(0, 0, tmxmove, tmymove)
    deltaangle = moveangle - lineangle

    If UGT(deltaangle, ANG180) Then deltaangle = deltaangle + ANG180
    '  I_Error ("SlideLine: ang>ANG180");

    lineangle = FINEANGLE(lineangle)
    deltaangle = FINEANGLE(deltaangle)

    movelen = P_AproxDistance(tmxmove, tmymove)
    newlen = FixedMul(movelen, finecosine(deltaangle))

    tmxmove = FixedMul(newlen, finecosine(lineangle))
    tmymove = FixedMul(newlen, finesine(lineangle))
End Sub

' PTR_SlideTraverse
Public Function PTR_SlideTraverse(ByVal inr As Long) As Boolean
    Dim li As Long

    If Not intercepts(inr).isaline Then I_Error "PTR_SlideTraverse: not a line?"

    li = intercepts(inr).dline

    If (lines(li).flags And ML_TWOSIDED) = 0 Then
        If P_PointOnLineSide(mobjs(slidemo).x, mobjs(slidemo).y, li) <> 0 Then
            ' don't hit the back side
            PTR_SlideTraverse = True
            Exit Function
        End If
        GoTo isblocking
    End If

    ' set openrange, opentop, openbottom
    P_LineOpening li

    If openrange < mobjs(slidemo).height Then GoTo isblocking ' doesn't fit

    If opentop - mobjs(slidemo).z < mobjs(slidemo).height Then GoTo isblocking ' mobj is too high

    If openbottom - mobjs(slidemo).z > 24 * FRACUNIT Then GoTo isblocking ' too big a step up

    ' this line doesn't block movement
    PTR_SlideTraverse = True
    Exit Function

    ' the line does block movement,
    ' see if it is closer than best so far
isblocking:
    If intercepts(inr).frac < bestslidefrac Then
        secondslidefrac = bestslidefrac
        secondslideline = bestslideline
        bestslidefrac = intercepts(inr).frac
        bestslideline = li
    End If

    PTR_SlideTraverse = False                    ' stop
End Function

' P_SlideMove - the momx / momy move is bad, so try to slide along a wall.
' Find the first line hit, move flush to it, and slide along it.
' This is a kludgy mess.
Public Sub P_SlideMove(ByVal mo As Long)
    Dim leadx As Long, leady As Long, trailx As Long, traily As Long
    Dim newx As Long, newy As Long, hitcount As Long

    slidemo = mo
    hitcount = 0

retry:
    hitcount = hitcount + 1
    If hitcount = 3 Then GoTo stairstep          ' don't loop forever

    ' trace along the three leading corners
    If mobjs(mo).momx > 0 Then
        leadx = mobjs(mo).x + mobjs(mo).radius
        trailx = mobjs(mo).x - mobjs(mo).radius
    Else
        leadx = mobjs(mo).x - mobjs(mo).radius
        trailx = mobjs(mo).x + mobjs(mo).radius
    End If

    If mobjs(mo).momy > 0 Then
        leady = mobjs(mo).y + mobjs(mo).radius
        traily = mobjs(mo).y - mobjs(mo).radius
    Else
        leady = mobjs(mo).y - mobjs(mo).radius
        traily = mobjs(mo).y + mobjs(mo).radius
    End If

    bestslidefrac = FRACUNIT + 1

    P_PathTraverse leadx, leady, leadx + mobjs(mo).momx, leady + mobjs(mo).momy, PT_ADDLINES, TRAV_SLIDE
    P_PathTraverse trailx, leady, trailx + mobjs(mo).momx, leady + mobjs(mo).momy, PT_ADDLINES, TRAV_SLIDE
    P_PathTraverse leadx, traily, leadx + mobjs(mo).momx, traily + mobjs(mo).momy, PT_ADDLINES, TRAV_SLIDE

    ' move up to the wall
    If bestslidefrac = FRACUNIT + 1 Then
        ' the move most have hit the middle, so stairstep
stairstep:
        If Not P_TryMove(mo, mobjs(mo).x, mobjs(mo).y + mobjs(mo).momy) Then
            P_TryMove mo, mobjs(mo).x + mobjs(mo).momx, mobjs(mo).y
        End If
        Exit Sub
    End If

    ' fudge a bit to make sure it doesn't hit
    bestslidefrac = bestslidefrac - &H800&
    If bestslidefrac > 0 Then
        newx = FixedMul(mobjs(mo).momx, bestslidefrac)
        newy = FixedMul(mobjs(mo).momy, bestslidefrac)

        If Not P_TryMove(mo, mobjs(mo).x + newx, mobjs(mo).y + newy) Then GoTo stairstep
    End If

    ' Now continue along the wall.
    ' First calculate remainder.
    bestslidefrac = FRACUNIT - (bestslidefrac + &H800&)

    If bestslidefrac > FRACUNIT Then bestslidefrac = FRACUNIT

    If bestslidefrac <= 0 Then Exit Sub

    tmxmove = FixedMul(mobjs(mo).momx, bestslidefrac)
    tmymove = FixedMul(mobjs(mo).momy, bestslidefrac)

    P_HitSlideLine bestslideline                 ' clip the moves

    mobjs(mo).momx = tmxmove
    mobjs(mo).momy = tmymove

    If Not P_TryMove(mo, mobjs(mo).x + tmxmove, mobjs(mo).y + tmymove) Then GoTo retry
End Sub

' PTR_AimTraverse - sets linetaget and aimslope when a target is aimed at.
Public Function PTR_AimTraverse(ByVal inr As Long) As Boolean
    Dim li As Long, th As Long, slope As Long, thingtopslope As Long
    Dim thingbottomslope As Long, dist As Long

    If intercepts(inr).isaline Then
        li = intercepts(inr).dline

        If (lines(li).flags And ML_TWOSIDED) = 0 Then
            PTR_AimTraverse = False              ' stop
            Exit Function
        End If

        ' Crosses a two sided line.
        ' A two sided line will restrict the possible target ranges.
        P_LineOpening li

        If openbottom >= opentop Then
            PTR_AimTraverse = False              ' stop
            Exit Function
        End If

        dist = FixedMul(attackrange, intercepts(inr).frac)

        If sectors(lines(li).frontsector).floorheight <> sectors(lines(li).backsector).floorheight Then
            slope = FixedDiv(openbottom - shootz, dist)
            If slope > bottomslope Then bottomslope = slope
        End If

        If sectors(lines(li).frontsector).ceilingheight <> sectors(lines(li).backsector).ceilingheight Then
            slope = FixedDiv(opentop - shootz, dist)
            If slope < topslope Then topslope = slope
        End If

        If topslope <= bottomslope Then
            PTR_AimTraverse = False              ' stop
            Exit Function
        End If

        PTR_AimTraverse = True                   ' shot continues
        Exit Function
    End If

    ' shoot a thing
    th = intercepts(inr).dthing
    If th = shootthing Then
        PTR_AimTraverse = True                   ' can't shoot self
        Exit Function
    End If

    If (mobjs(th).flags And MF_SHOOTABLE) = 0 Then
        PTR_AimTraverse = True                   ' corpse or something
        Exit Function
    End If

    ' check angles to see if the thing can be aimed at
    dist = FixedMul(attackrange, intercepts(inr).frac)
    thingtopslope = FixedDiv(mobjs(th).z + mobjs(th).height - shootz, dist)

    If thingtopslope < bottomslope Then
        PTR_AimTraverse = True                   ' shot over the thing
        Exit Function
    End If

    thingbottomslope = FixedDiv(mobjs(th).z - shootz, dist)

    If thingbottomslope > topslope Then
        PTR_AimTraverse = True                   ' shot under the thing
        Exit Function
    End If

    ' this thing can be hit!
    If thingtopslope > topslope Then thingtopslope = topslope

    If thingbottomslope < bottomslope Then thingbottomslope = bottomslope

    aimslope = (thingtopslope + thingbottomslope) \ 2
    linetarget = th

    PTR_AimTraverse = False                      ' don't go any farther
End Function

' PTR_ShootTraverse
Public Function PTR_ShootTraverse(ByVal inr As Long) As Boolean
    Dim x As Long, y As Long, z As Long, frac As Long, li As Long, th As Long
    Dim slope As Long, dist As Long, thingtopslope As Long, thingbottomslope As Long

    If intercepts(inr).isaline Then
        li = intercepts(inr).dline

        If lines(li).special <> 0 Then P_ShootSpecialLine shootthing, li

        If (lines(li).flags And ML_TWOSIDED) = 0 Then GoTo hitline

        ' crosses a two sided line
        P_LineOpening li

        dist = FixedMul(attackrange, intercepts(inr).frac)

        If sectors(lines(li).frontsector).floorheight <> sectors(lines(li).backsector).floorheight Then
            slope = FixedDiv(openbottom - shootz, dist)
            If slope > aimslope Then GoTo hitline
        End If

        If sectors(lines(li).frontsector).ceilingheight <> sectors(lines(li).backsector).ceilingheight Then
            slope = FixedDiv(opentop - shootz, dist)
            If slope < aimslope Then GoTo hitline
        End If

        ' shot continues
        PTR_ShootTraverse = True
        Exit Function

        ' hit line
hitline:
        ' position a bit closer
        frac = intercepts(inr).frac - FixedDiv(4 * FRACUNIT, attackrange)
        x = trace.x + FixedMul(trace.dx, frac)
        y = trace.y + FixedMul(trace.dy, frac)
        z = shootz + FixedMul(aimslope, FixedMul(frac, attackrange))

        If sectors(lines(li).frontsector).ceilingpic = skyflatnum Then
            ' don't shoot the sky!
            If z > sectors(lines(li).frontsector).ceilingheight Then
                PTR_ShootTraverse = False
                Exit Function
            End If

            ' it's a sky hack wall
            If lines(li).backsector <> -1 Then
                If sectors(lines(li).backsector).ceilingpic = skyflatnum Then
                    PTR_ShootTraverse = False
                    Exit Function
                End If
            End If
        End If

        ' Spawn bullet puffs.
        P_SpawnPuff x, y, z

        ' don't go any farther
        PTR_ShootTraverse = False
        Exit Function
    End If

    ' shoot a thing
    th = intercepts(inr).dthing
    If th = shootthing Then
        PTR_ShootTraverse = True                 ' can't shoot self
        Exit Function
    End If

    If (mobjs(th).flags And MF_SHOOTABLE) = 0 Then
        PTR_ShootTraverse = True                 ' corpse or something
        Exit Function
    End If

    ' check angles to see if the thing can be aimed at
    dist = FixedMul(attackrange, intercepts(inr).frac)
    thingtopslope = FixedDiv(mobjs(th).z + mobjs(th).height - shootz, dist)

    If thingtopslope < aimslope Then
        PTR_ShootTraverse = True                 ' shot over the thing
        Exit Function
    End If

    thingbottomslope = FixedDiv(mobjs(th).z - shootz, dist)

    If thingbottomslope > aimslope Then
        PTR_ShootTraverse = True                 ' shot under the thing
        Exit Function
    End If

    ' hit thing
    ' position a bit closer
    frac = intercepts(inr).frac - FixedDiv(10 * FRACUNIT, attackrange)

    x = trace.x + FixedMul(trace.dx, frac)
    y = trace.y + FixedMul(trace.dy, frac)
    z = shootz + FixedMul(aimslope, FixedMul(frac, attackrange))

    ' Spawn bullet puffs or blod spots, depending on target type.
    If (mobjs(th).flags And MF_NOBLOOD) <> 0 Then
        P_SpawnPuff x, y, z
    Else
        P_SpawnBlood x, y, z, la_damage
    End If

    If la_damage <> 0 Then P_DamageMobj th, shootthing, shootthing, la_damage

    ' don't go any farther
    PTR_ShootTraverse = False
End Function

' P_AimLineAttack
Public Function P_AimLineAttack(ByVal t1 As Long, ByVal angle As Long, ByVal distance As Long) As Long
    Dim x2 As Long, y2 As Long

    angle = FINEANGLE(angle)
    shootthing = t1

    x2 = mobjs(t1).x + FIXTOINT(distance) * finecosine(angle)
    y2 = mobjs(t1).y + FIXTOINT(distance) * finesine(angle)
    shootz = mobjs(t1).z + SAR(mobjs(t1).height, 1) + 8 * FRACUNIT

    ' can't shoot outside view angles
    topslope = 100 * FRACUNIT \ 160
    bottomslope = -100 * FRACUNIT \ 160

    attackrange = distance
    linetarget = 0

    P_PathTraverse mobjs(t1).x, mobjs(t1).y, x2, y2, PT_ADDLINES Or PT_ADDTHINGS, TRAV_AIM

    If linetarget <> 0 Then
        P_AimLineAttack = aimslope
    Else
        P_AimLineAttack = 0
    End If
End Function

' P_LineAttack - if damage == 0, it is just a test trace that will leave
' linetarget set.
Public Sub P_LineAttack(ByVal t1 As Long, ByVal angle As Long, ByVal distance As Long, ByVal slope As Long, ByVal damage As Long)
    Dim x2 As Long, y2 As Long

    angle = FINEANGLE(angle)
    shootthing = t1
    la_damage = damage
    x2 = mobjs(t1).x + FIXTOINT(distance) * finecosine(angle)
    y2 = mobjs(t1).y + FIXTOINT(distance) * finesine(angle)
    shootz = mobjs(t1).z + SAR(mobjs(t1).height, 1) + 8 * FRACUNIT
    attackrange = distance
    aimslope = slope

    P_PathTraverse mobjs(t1).x, mobjs(t1).y, x2, y2, PT_ADDLINES Or PT_ADDTHINGS, TRAV_SHOOT
End Sub

' USE LINES

Public Function PTR_UseTraverse(ByVal inr As Long) As Boolean
    Dim side As Long, li As Long

    li = intercepts(inr).dline
    If lines(li).special = 0 Then
        P_LineOpening li
        If openrange <= 0 Then
            S_StartSound usething, sfx_noway

            ' can't use through a wall
            PTR_UseTraverse = False
            Exit Function
        End If
        ' not a special line, but keep checking
        PTR_UseTraverse = True
        Exit Function
    End If

    side = 0
    If P_PointOnLineSide(mobjs(usething).x, mobjs(usething).y, li) = 1 Then side = 1

    '  return false;           // don't use back side

    P_UseSpecialLine usething, li, side

    ' can't use for than one special line in a row
    PTR_UseTraverse = False
End Function

' P_UseLines - looks for special lines in front of the player to activate.
Public Sub P_UseLines(ByVal player As Long)
    Dim angle As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, mo As Long

    mo = players(player).mo
    usething = mo

    angle = FINEANGLE(mobjs(mo).angle)

    x1 = mobjs(mo).x
    y1 = mobjs(mo).y
    x2 = x1 + (USERANGE \ FRACUNIT) * finecosine(angle)
    y2 = y1 + (USERANGE \ FRACUNIT) * finesine(angle)

    P_PathTraverse x1, y1, x2, y2, PT_ADDLINES, TRAV_USE
End Sub

' RADIUS ATTACK

' PIT_RadiusAttack - "bombsource" is the creature that caused the explosion
' at "bombspot".
Public Function PIT_RadiusAttack(ByVal thing As Long) As Boolean
    Dim dx As Long, dy As Long, dist As Long

    PIT_RadiusAttack = True

    If (mobjs(thing).flags And MF_SHOOTABLE) = 0 Then Exit Function

    ' Boss spider and cyborg take no damage from concussion.
    If mobjs(thing).mtype = MT_CYBORG Or mobjs(thing).mtype = MT_SPIDER Then Exit Function

    dx = Abs(mobjs(thing).x - mobjs(bombspot).x)
    dy = Abs(mobjs(thing).y - mobjs(bombspot).y)

    If dx > dy Then dist = dx Else dist = dy
    dist = SAR(dist - mobjs(thing).radius, FRACBITS)

    If dist < 0 Then dist = 0

    If dist >= bombdamage Then Exit Function     ' out of range

    If P_CheckSight(thing, bombspot) Then
        ' must be in direct path
        P_DamageMobj thing, bombspot, bombsource, bombdamage - dist
    End If
End Function

' P_RadiusAttack - source is the creature that caused the explosion at spot.
Public Sub P_RadiusAttack(ByVal spot As Long, ByVal source As Long, ByVal damage As Long)
    Dim x As Long, y As Long, xl As Long, xh As Long, yl As Long, yh As Long, dist As Long

    dist = (damage + MAXRADIUS) * FRACUNIT       ' (sic: wraps exactly like the C code)
    yh = SAR(mobjs(spot).y + dist - bmaporgy, MAPBLOCKSHIFT)
    yl = SAR(mobjs(spot).y - dist - bmaporgy, MAPBLOCKSHIFT)
    xh = SAR(mobjs(spot).x + dist - bmaporgx, MAPBLOCKSHIFT)
    xl = SAR(mobjs(spot).x - dist - bmaporgx, MAPBLOCKSHIFT)
    bombspot = spot
    bombsource = source
    bombdamage = damage

    For y = yl To yh
        For x = xl To xh
            P_BlockThingsIterator x, y, ITER_RADIUSATTACK
        Next
    Next
End Sub

' SECTOR HEIGHT CHANGING
' After modifying a sectors floor or ceiling height, call this routine to
' adjust the positions of all things that touch the sector.
'
' If anything doesn't fit anymore, true will be returned.
' If crunch is true, they will take damage as they are being crushed.
' If Crunch is false, you should set the sector height back the way it was
' and call P_ChangeSector again to undo the changes.

' PIT_ChangeSector
Public Function PIT_ChangeSector(ByVal thing As Long) As Boolean
    Dim mo As Long

    PIT_ChangeSector = True

    If P_ThingHeightClip(thing) Then Exit Function ' keep checking

    ' crunch bodies to giblets
    If mobjs(thing).health <= 0 Then
        P_SetMobjState thing, S_GIBS

        mobjs(thing).flags = mobjs(thing).flags And (Not MF_SOLID)
        mobjs(thing).height = 0
        mobjs(thing).radius = 0

        ' keep checking
        Exit Function
    End If

    ' crunch dropped items
    If (mobjs(thing).flags And MF_DROPPED) <> 0 Then
        P_RemoveMobj thing

        ' keep checking
        Exit Function
    End If

    If (mobjs(thing).flags And MF_SHOOTABLE) = 0 Then
        ' assume it is bloody gibs or something
        Exit Function
    End If

    nofit = True

    If crushchange And (leveltime And 3) = 0 Then
        P_DamageMobj thing, 0, 0, 10

        ' spray blood in a random direction
        mo = P_SpawnMobj(mobjs(thing).x, mobjs(thing).y, mobjs(thing).z + mobjs(thing).height \ 2, MT_BLOOD)

        mobjs(mo).momx = P_SubRandom() * 4096
        mobjs(mo).momy = P_SubRandom() * 4096
    End If

    ' keep checking (crush other things)
End Function

' P_ChangeSector
Public Function P_ChangeSector(ByVal sector As Long, ByVal crunch As Boolean) As Boolean
    Dim x As Long, y As Long

    nofit = False
    crushchange = crunch

    ' re-check heights for all things near the moving sector
    For x = sectors(sector).blockbox(BOXLEFT) To sectors(sector).blockbox(BOXRIGHT)
        For y = sectors(sector).blockbox(BOXBOTTOM) To sectors(sector).blockbox(BOXTOP)
            P_BlockThingsIterator x, y, ITER_CHANGESECTOR
        Next
    Next

    P_ChangeSector = nofit
End Function
