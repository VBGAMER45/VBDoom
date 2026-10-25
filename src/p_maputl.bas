Attribute VB_Name = "p_maputl"
' p_maputl.bas - movement/collision utility functions, as used by function
' in p_map.c.  BLOCKMAP Iterator functions, and some PIT_* functions to use
' for iteration (p_maputl.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private earlyout As Boolean
Private maxintercepts As Long

' P_AproxDistance - gives an estimation of distance (not exact)
Public Function P_AproxDistance(ByVal dx As Long, ByVal dy As Long) As Long
    dx = Abs(dx)
    dy = Abs(dy)
    If dx < dy Then
        P_AproxDistance = dx + dy - SAR(dx, 1)
    Else
        P_AproxDistance = dx + dy - SAR(dy, 1)
    End If
End Function

' P_PointOnLineSide - returns 0 or 1
Public Function P_PointOnLineSide(ByVal x As Long, ByVal y As Long, ByVal line As Long) As Long
    Dim dx As Long, dy As Long, lft As Long, rgt As Long, ldx As Long, ldy As Long

    ldx = lines(line).dx
    ldy = lines(line).dy
    If ldx = 0 Then
        If x <= vertexes(lines(line).v1).x Then
            If ldy > 0 Then P_PointOnLineSide = 1 Else P_PointOnLineSide = 0
        Else
            If ldy < 0 Then P_PointOnLineSide = 1 Else P_PointOnLineSide = 0
        End If
        Exit Function
    End If
    If ldy = 0 Then
        If y <= vertexes(lines(line).v1).y Then
            If ldx < 0 Then P_PointOnLineSide = 1 Else P_PointOnLineSide = 0
        Else
            If ldx > 0 Then P_PointOnLineSide = 1 Else P_PointOnLineSide = 0
        End If
        Exit Function
    End If

    dx = x - vertexes(lines(line).v1).x
    dy = y - vertexes(lines(line).v1).y

    lft = FixedMul((ldy And &HFFFF0000) \ &H10000, dx)
    rgt = FixedMul(dy, (ldx And &HFFFF0000) \ &H10000)

    If rgt < lft Then
        P_PointOnLineSide = 0                    ' front side
    Else
        P_PointOnLineSide = 1                    ' back side
    End If
End Function

' P_BoxOnLineSide - considers the line to be infinite.  Returns side 0 or 1,
' -1 if box crosses the line.
Public Function P_BoxOnLineSide(tmbox() As Long, ByVal ld As Long) As Long
    Dim p1 As Long, p2 As Long, v1x As Long, v1y As Long

    v1x = vertexes(lines(ld).v1).x
    v1y = vertexes(lines(ld).v1).y

    Select Case lines(ld).slopetype
        Case ST_HORIZONTAL
            If tmbox(BOXTOP) > v1y Then p1 = 1 Else p1 = 0
            If tmbox(BOXBOTTOM) > v1y Then p2 = 1 Else p2 = 0
            If lines(ld).dx < 0 Then
                p1 = p1 Xor 1
                p2 = p2 Xor 1
            End If
        Case ST_VERTICAL
            If tmbox(BOXRIGHT) < v1x Then p1 = 1 Else p1 = 0
            If tmbox(BOXLEFT) < v1x Then p2 = 1 Else p2 = 0
            If lines(ld).dy < 0 Then
                p1 = p1 Xor 1
                p2 = p2 Xor 1
            End If
        Case ST_POSITIVE
            p1 = P_PointOnLineSide(tmbox(BOXLEFT), tmbox(BOXTOP), ld)
            p2 = P_PointOnLineSide(tmbox(BOXRIGHT), tmbox(BOXBOTTOM), ld)
        Case ST_NEGATIVE
            p1 = P_PointOnLineSide(tmbox(BOXRIGHT), tmbox(BOXTOP), ld)
            p2 = P_PointOnLineSide(tmbox(BOXLEFT), tmbox(BOXBOTTOM), ld)
    End Select

    If p1 = p2 Then
        P_BoxOnLineSide = p1
    Else
        P_BoxOnLineSide = -1
    End If
End Function

' P_PointOnDivlineSide - returns 0 or 1.
Public Function P_PointOnDivlineSide(ByVal x As Long, ByVal y As Long, dl As divline_t) As Long
    Dim dx As Long, dy As Long, lft As Long, rgt As Long

    If dl.dx = 0 Then
        If x <= dl.x Then
            If dl.dy > 0 Then P_PointOnDivlineSide = 1 Else P_PointOnDivlineSide = 0
        Else
            If dl.dy < 0 Then P_PointOnDivlineSide = 1 Else P_PointOnDivlineSide = 0
        End If
        Exit Function
    End If
    If dl.dy = 0 Then
        If y <= dl.y Then
            If dl.dx < 0 Then P_PointOnDivlineSide = 1 Else P_PointOnDivlineSide = 0
        Else
            If dl.dx > 0 Then P_PointOnDivlineSide = 1 Else P_PointOnDivlineSide = 0
        End If
        Exit Function
    End If

    dx = x - dl.x
    dy = y - dl.y

    ' try to quickly decide by looking at sign bits
    If (dl.dy Xor dl.dx Xor dx Xor dy) < 0 Then
        If (dl.dy Xor dx) < 0 Then
            P_PointOnDivlineSide = 1             ' (left is negative)
        Else
            P_PointOnDivlineSide = 0
        End If
        Exit Function
    End If

    lft = FixedMul(SAR(dl.dy, 8), SAR(dx, 8))
    rgt = FixedMul(SAR(dy, 8), SAR(dl.dx, 8))

    If rgt < lft Then
        P_PointOnDivlineSide = 0                 ' front side
    Else
        P_PointOnDivlineSide = 1                 ' back side
    End If
End Function

' P_MakeDivline
Public Sub P_MakeDivline(ByVal li As Long, dl As divline_t)
    dl.x = vertexes(lines(li).v1).x
    dl.y = vertexes(lines(li).v1).y
    dl.dx = lines(li).dx
    dl.dy = lines(li).dy
End Sub

' P_InterceptVector - returns the fractional intercept point along the first
' divline.  This is only called by the addthings and addlines traversers.
Public Function P_InterceptVector(v2 As divline_t, v1 As divline_t) As Long
    Dim num As Long, den As Long

    den = FixedMul(SAR(v1.dy, 8), v2.dx) - FixedMul(SAR(v1.dx, 8), v2.dy)

    If den = 0 Then
        P_InterceptVector = 0
        Exit Function
    End If

    num = FixedMul(SAR(v1.x - v2.x, 8), v1.dy) + FixedMul(SAR(v2.y - v1.y, 8), v1.dx)

    P_InterceptVector = FixedDiv(num, den)
End Function

' P_LineOpening - sets opentop and openbottom to the window through a two
' sided line.  OPTIMIZE: keep this precalculated
Public Sub P_LineOpening(ByVal linedef As Long)
    Dim front As Long, back As Long

    If lines(linedef).sidenum(1) = -1 Then
        ' single sided line
        openrange = 0
        Exit Sub
    End If

    front = lines(linedef).frontsector
    back = lines(linedef).backsector

    If sectors(front).ceilingheight < sectors(back).ceilingheight Then
        opentop = sectors(front).ceilingheight
    Else
        opentop = sectors(back).ceilingheight
    End If

    If sectors(front).floorheight > sectors(back).floorheight Then
        openbottom = sectors(front).floorheight
        lowfloor = sectors(back).floorheight
    Else
        openbottom = sectors(back).floorheight
        lowfloor = sectors(front).floorheight
    End If

    openrange = opentop - openbottom
End Sub

' THING POSITION SETTING

' P_UnsetThingPosition - unlinks a thing from block map and sectors.  On each
' position change, BLOCKMAP and other lookups maintaining lists ot things
' inside these structures need to be updated.
Public Sub P_UnsetThingPosition(ByVal thing As Long)
    Dim blockx As Long, blocky As Long, sn As Long, sp As Long, bn As Long, bp As Long, t As Long

    If checklinks And (mobjs(thing).flags And MF_NOSECTOR) = 0 Then
        t = sectors(subsectors(mobjs(thing).subsector).sector).thinglist
        Do While t <> 0
            If t = thing Then Exit Do
            t = mobjs(t).snext
        Loop
        If t = 0 Then
            Err.Raise vbObjectError + 1, "P_UnsetThingPosition", "P_UnsetThingPosition: mobj " & thing & " (type " & mobjs(thing).mtype & ", func " & thinkers(thing).func & _
                    ") not in the list of its sector " & subsectors(mobjs(thing).subsector).sector & _
                    "; running thinker " & currentThinker & " type " & mobjs(currentThinker).mtype & " at gametic " & gametic & _
                    "; called from " & dbgUnsetBy
        End If
    End If

    If (mobjs(thing).flags And MF_NOSECTOR) = 0 Then
        ' inert things don't need to be in blockmap?
        ' unlink from subsector
        sn = mobjs(thing).snext
        sp = mobjs(thing).sprev
        If sn <> 0 Then mobjs(sn).sprev = sp

        If sp <> 0 Then
            mobjs(sp).snext = sn
        Else
            sectors(subsectors(mobjs(thing).subsector).sector).thinglist = sn
        End If
    End If

    If (mobjs(thing).flags And MF_NOBLOCKMAP) = 0 Then
        ' inert things don't need to be in blockmap
        ' unlink from block map
        bn = mobjs(thing).bnext
        bp = mobjs(thing).bprev
        If bn <> 0 Then mobjs(bn).bprev = bp

        If bp <> 0 Then
            mobjs(bp).bnext = bn
        Else
            blockx = SAR(mobjs(thing).x - bmaporgx, MAPBLOCKSHIFT)
            blocky = SAR(mobjs(thing).y - bmaporgy, MAPBLOCKSHIFT)

            If blockx >= 0 And blockx < bmapwidth And blocky >= 0 And blocky < bmapheight Then
                blocklinks(blocky * bmapwidth + blockx) = bn
            End If
        End If
    End If
End Sub

' P_SetThingPosition - links a thing into both a block and a subsector based
' on it's x y.  Sets thing->subsector properly
Public Sub P_SetThingPosition(ByVal thing As Long)
    Dim ss As Long, sec As Long, blockx As Long, blocky As Long, lnk As Long

    ' link into subsector
    ss = R_PointInSubsector(mobjs(thing).x, mobjs(thing).y)
    mobjs(thing).subsector = ss

    If (mobjs(thing).flags And MF_NOSECTOR) = 0 Then
        ' invisible things don't go into the sector links
        sec = subsectors(ss).sector

        mobjs(thing).sprev = 0
        mobjs(thing).snext = sectors(sec).thinglist

        If sectors(sec).thinglist <> 0 Then mobjs(sectors(sec).thinglist).sprev = thing

        sectors(sec).thinglist = thing
    End If

    ' link into blockmap
    If (mobjs(thing).flags And MF_NOBLOCKMAP) = 0 Then
        ' inert things don't need to be in blockmap
        blockx = SAR(mobjs(thing).x - bmaporgx, MAPBLOCKSHIFT)
        blocky = SAR(mobjs(thing).y - bmaporgy, MAPBLOCKSHIFT)

        If blockx >= 0 And blockx < bmapwidth And blocky >= 0 And blocky < bmapheight Then
            lnk = blocky * bmapwidth + blockx
            mobjs(thing).bprev = 0
            mobjs(thing).bnext = blocklinks(lnk)
            If blocklinks(lnk) <> 0 Then mobjs(blocklinks(lnk)).bprev = thing

            blocklinks(lnk) = thing
        Else
            ' thing is off the map
            mobjs(thing).bnext = 0
            mobjs(thing).bprev = 0
        End If
    End If
End Sub

' BLOCK MAP ITERATORS
' For each line/thing in the given mapblock, call the passed PIT_* function.
' If the function returns false, exit with false without checking anything
' else.

' P_BlockLinesIterator - the validcount flags are used to avoid checking
' lines that are marked in multiple mapblocks, so increment validcount
' before the first call to P_BlockLinesIterator, then make one or more calls
' to it.
Public Function P_BlockLinesIterator(ByVal x As Long, ByVal y As Long, ByVal func As Long) As Boolean
    Dim offset As Long, ld As Long, ok As Boolean

    If x < 0 Or y < 0 Or x >= bmapwidth Or y >= bmapheight Then
        P_BlockLinesIterator = True
        Exit Function
    End If

    offset = blockmaplump(blockmap + y * bmapwidth + x) And &HFFFF&

    ld = blockmaplump(offset)
    Do While ld <> -1
        If lines(ld).validcount <> validcount Then
            lines(ld).validcount = validcount     ' mark the line as checked

            Select Case func
                Case ITER_CHECKLINE: ok = PIT_CheckLine(ld)
                Case ITER_ADDLINEINTERCEPTS: ok = PIT_AddLineIntercepts(ld)
                Case Else: I_Error "P_BlockLinesIterator: bad iterator " & func
            End Select
            If Not ok Then
                P_BlockLinesIterator = False
                Exit Function
            End If
        End If
        offset = offset + 1
        ld = blockmaplump(offset)
    Loop
    P_BlockLinesIterator = True                  ' everything was checked
End Function

' P_BlockThingsIterator
Public Function P_BlockThingsIterator(ByVal x As Long, ByVal y As Long, ByVal func As Long) As Boolean
    Dim mobj As Long, ok As Boolean

    If x < 0 Or y < 0 Or x >= bmapwidth Or y >= bmapheight Then
        P_BlockThingsIterator = True
        Exit Function
    End If

    mobj = blocklinks(y * bmapwidth + x)
    Do While mobj <> 0
        Select Case func
            Case ITER_CHECKTHING: ok = PIT_CheckThing(mobj)
            Case ITER_STOMPTHING: ok = PIT_StompThing(mobj)
            Case ITER_CHANGESECTOR: ok = PIT_ChangeSector(mobj)
            Case ITER_RADIUSATTACK: ok = PIT_RadiusAttack(mobj)
            Case ITER_VILECHECK: ok = PIT_VileCheck(mobj)
            Case ITER_ADDTHINGINTERCEPTS: ok = PIT_AddThingIntercepts(mobj)
            Case Else: I_Error "P_BlockThingsIterator: bad iterator " & func
        End Select
        If Not ok Then
            P_BlockThingsIterator = False
            Exit Function
        End If
        mobj = mobjs(mobj).bnext
    Loop
    P_BlockThingsIterator = True
End Function

' INTERCEPT ROUTINES

Private Sub NewIntercept(ByVal frac As Long, ByVal isaline As Boolean, ByVal thing As Long, ByVal line As Long)
    If intercept_p >= maxintercepts Then
        maxintercepts = maxintercepts * 2
        ReDim Preserve intercepts(0 To maxintercepts - 1)
    End If
    intercepts(intercept_p).frac = frac
    intercepts(intercept_p).isaline = isaline
    intercepts(intercept_p).dthing = thing
    intercepts(intercept_p).dline = line
    intercept_p = intercept_p + 1
End Sub

' PIT_AddLineIntercepts.  Looks for lines in the given block that intercept
' the given trace to add to the intercepts list.  A line is crossed if its
' endpoints are on opposite sides of the trace.  Returns true if earlyout and
' a solid line hit.
Private Function PIT_AddLineIntercepts(ByVal ld As Long) As Boolean
    Dim s1 As Long, s2 As Long, frac As Long, dl As divline_t

    ' avoid precision problems with two routines
    If trace.dx > FRACUNIT * 16 Or trace.dy > FRACUNIT * 16 Or _
       trace.dx < -FRACUNIT * 16 Or trace.dy < -FRACUNIT * 16 Then
        s1 = P_PointOnDivlineSide(vertexes(lines(ld).v1).x, vertexes(lines(ld).v1).y, trace)
        s2 = P_PointOnDivlineSide(vertexes(lines(ld).v2).x, vertexes(lines(ld).v2).y, trace)
    Else
        s1 = P_PointOnLineSide(trace.x, trace.y, ld)
        s2 = P_PointOnLineSide(trace.x + trace.dx, trace.y + trace.dy, ld)
    End If

    If s1 = s2 Then
        PIT_AddLineIntercepts = True             ' line isn't crossed
        Exit Function
    End If

    ' hit the line
    P_MakeDivline ld, dl
    frac = P_InterceptVector(trace, dl)

    If frac < 0 Then
        PIT_AddLineIntercepts = True             ' behind source
        Exit Function
    End If

    ' try to early out the check
    If earlyout And frac < FRACUNIT And lines(ld).backsector = -1 Then
        PIT_AddLineIntercepts = False            ' stop checking
        Exit Function
    End If

    NewIntercept frac, True, 0, ld

    PIT_AddLineIntercepts = True                 ' continue
End Function

' PIT_AddThingIntercepts
Private Function PIT_AddThingIntercepts(ByVal thing As Long) As Boolean
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long, s1 As Long, s2 As Long
    Dim tracepositive As Boolean, dl As divline_t, frac As Long, r As Long

    tracepositive = ((trace.dx Xor trace.dy) > 0)

    r = mobjs(thing).radius
    ' check a corner to corner crossection for hit
    If tracepositive Then
        x1 = mobjs(thing).x - r
        y1 = mobjs(thing).y + r
        x2 = mobjs(thing).x + r
        y2 = mobjs(thing).y - r
    Else
        x1 = mobjs(thing).x - r
        y1 = mobjs(thing).y - r
        x2 = mobjs(thing).x + r
        y2 = mobjs(thing).y + r
    End If

    s1 = P_PointOnDivlineSide(x1, y1, trace)
    s2 = P_PointOnDivlineSide(x2, y2, trace)

    If s1 = s2 Then
        PIT_AddThingIntercepts = True            ' line isn't crossed
        Exit Function
    End If

    dl.x = x1
    dl.y = y1
    dl.dx = x2 - x1
    dl.dy = y2 - y1

    frac = P_InterceptVector(trace, dl)

    If frac < 0 Then
        PIT_AddThingIntercepts = True            ' behind source
        Exit Function
    End If

    NewIntercept frac, False, thing, -1

    PIT_AddThingIntercepts = True                ' keep going
End Function

' P_TraverseIntercepts - returns true if the traverser function returns
' true for all lines.
Private Function P_TraverseIntercepts(ByVal func As Long, ByVal maxfrac As Long) As Boolean
    Dim count As Long, dist As Long, scan As Long, inr As Long, ok As Boolean

    count = intercept_p
    inr = -1                                     ' shut up compiler warning

    Do While count > 0
        count = count - 1
        dist = MAXINT
        For scan = 0 To intercept_p - 1
            If intercepts(scan).frac < dist Then
                dist = intercepts(scan).frac
                inr = scan
            End If
        Next

        If dist > maxfrac Then
            P_TraverseIntercepts = True          ' checked everything in range
            Exit Function
        End If

        Select Case func
            Case TRAV_SLIDE: ok = PTR_SlideTraverse(inr)
            Case TRAV_AIM: ok = PTR_AimTraverse(inr)
            Case TRAV_SHOOT: ok = PTR_ShootTraverse(inr)
            Case TRAV_USE: ok = PTR_UseTraverse(inr)
            Case Else: I_Error "P_TraverseIntercepts: bad traverser " & func
        End Select
        If Not ok Then
            P_TraverseIntercepts = False         ' don't bother going farther
            Exit Function
        End If

        intercepts(inr).frac = MAXINT
    Loop

    P_TraverseIntercepts = True                  ' everything was traversed
End Function

' P_PathTraverse - traces a line from x1,y1 to x2,y2, calling the traverser
' function for each.  Returns true if the traverser function returns true for
' all lines.
Public Function P_PathTraverse(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, _
                               ByVal flags As Long, ByVal trav As Long) As Boolean
    Dim xt1 As Long, yt1 As Long, xt2 As Long, yt2 As Long
    Dim xstep As Long, ystep As Long, partial As Long
    Dim xintercept As Long, yintercept As Long
    Dim mapx As Long, mapy As Long, mapxstep As Long, mapystep As Long, count As Long

    If maxintercepts = 0 Then
        maxintercepts = 256
        ReDim intercepts(0 To maxintercepts - 1)
    End If

    earlyout = ((flags And PT_EARLYOUT) <> 0)

    validcount = validcount + 1
    intercept_p = 0

    If ((x1 - bmaporgx) And MAPBMASK) = 0 Then x1 = x1 + FRACUNIT ' don't side exactly on a line

    If ((y1 - bmaporgy) And MAPBMASK) = 0 Then y1 = y1 + FRACUNIT ' don't side exactly on a line

    trace.x = x1
    trace.y = y1
    trace.dx = x2 - x1
    trace.dy = y2 - y1

    x1 = x1 - bmaporgx
    y1 = y1 - bmaporgy
    xt1 = SAR(x1, MAPBLOCKSHIFT)
    yt1 = SAR(y1, MAPBLOCKSHIFT)

    x2 = x2 - bmaporgx
    y2 = y2 - bmaporgy
    xt2 = SAR(x2, MAPBLOCKSHIFT)
    yt2 = SAR(y2, MAPBLOCKSHIFT)

    If xt2 > xt1 Then
        mapxstep = 1
        partial = FRACUNIT - (SAR(x1, MAPBTOFRAC) And (FRACUNIT - 1))
        ystep = FixedDiv(y2 - y1, Abs(x2 - x1))
    ElseIf xt2 < xt1 Then
        mapxstep = -1
        partial = SAR(x1, MAPBTOFRAC) And (FRACUNIT - 1)
        ystep = FixedDiv(y2 - y1, Abs(x2 - x1))
    Else
        mapxstep = 0
        partial = FRACUNIT
        ystep = 256 * FRACUNIT
    End If

    yintercept = SAR(y1, MAPBTOFRAC) + FixedMul(partial, ystep)

    If yt2 > yt1 Then
        mapystep = 1
        partial = FRACUNIT - (SAR(y1, MAPBTOFRAC) And (FRACUNIT - 1))
        xstep = FixedDiv(x2 - x1, Abs(y2 - y1))
    ElseIf yt2 < yt1 Then
        mapystep = -1
        partial = SAR(y1, MAPBTOFRAC) And (FRACUNIT - 1)
        xstep = FixedDiv(x2 - x1, Abs(y2 - y1))
    Else
        mapystep = 0
        partial = FRACUNIT
        xstep = 256 * FRACUNIT
    End If
    xintercept = SAR(x1, MAPBTOFRAC) + FixedMul(partial, xstep)

    ' Step through map blocks.
    ' Count is present to prevent a round off error from skipping the break.
    mapx = xt1
    mapy = yt1

    For count = 0 To 63
        If (flags And PT_ADDLINES) <> 0 Then
            If Not P_BlockLinesIterator(mapx, mapy, ITER_ADDLINEINTERCEPTS) Then
                P_PathTraverse = False           ' early out
                Exit Function
            End If
        End If

        If (flags And PT_ADDTHINGS) <> 0 Then
            If Not P_BlockThingsIterator(mapx, mapy, ITER_ADDTHINGINTERCEPTS) Then
                P_PathTraverse = False           ' early out
                Exit Function
            End If
        End If

        If mapx = xt2 And mapy = yt2 Then Exit For

        If ((yintercept And &HFFFF0000) \ &H10000) = mapy Then
            yintercept = yintercept + ystep
            mapx = mapx + mapxstep
        ElseIf ((xintercept And &HFFFF0000) \ &H10000) = mapx Then
            xintercept = xintercept + xstep
            mapy = mapy + mapystep
        End If
    Next

    ' go through the sorted list
    P_PathTraverse = P_TraverseIntercepts(trav, FRACUNIT)
End Function
