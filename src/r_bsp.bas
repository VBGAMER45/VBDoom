Attribute VB_Name = "r_bsp"
' r_bsp.bas - BSP traversal, handling of LineSegs for rendering (r_bsp.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' A drawseg_t; clip arrays are offsets into openings() (-1 = NULL).
Public Type drawseg_t
    curline As Long                              ' seg index
    x1 As Long
    x2 As Long
    scale1 As Long
    scale2 As Long
    scalestep As Long
    silhouette As Long                           ' 0=none, 1=bottom, 2=top, 3=both
    bsilheight As Long                           ' do not clip sprites above this
    tsilheight As Long                           ' do not clip sprites below this
    sprtopclip As Long                           ' openings() offset, -1 = NULL
    sprbottomclip As Long
    maskedtexturecol As Long
End Type

Public curline As Long
Public sidedef As Long
Public linedef As Long
Public frontsector As Long
Public backsector As Long                        ' -1 = none

Public drawsegs() As drawseg_t
Public ds_p As Long                              ' number of drawsegs in use (C: ds_p - drawsegs)
Private maxdrawsegs As Long

' solidsegs: a set of ranges of columns that are already drawn
Private Const MAXSEGS As Long = 400              ' C: 32 (overflow corrupted memory)
Private ss_first(0 To 400) As Long
Private ss_last(0 To 400) As Long
Private newend As Long

Private checkcoord(0 To 11, 0 To 3) As Long
Private checkcoordInit As Boolean

Public Sub R_ClearDrawSegs()
    If maxdrawsegs = 0 Then
        maxdrawsegs = 512
        ReDim drawsegs(0 To maxdrawsegs - 1)
    End If
    ds_p = 0
End Sub

' called by R_StoreWallRange before it fills drawsegs(ds_p)
Public Sub R_CheckDrawSegs()
    If ds_p >= maxdrawsegs Then
        If maxdrawsegs >= 65536 Then I_Error "R_CheckDrawSegs: runaway drawseg count"
        maxdrawsegs = maxdrawsegs * 2
        ReDim Preserve drawsegs(0 To maxdrawsegs - 1)
    End If
End Sub

' R_ClipSolidWallSegment - does handle solid walls, e.g. single sided LineDefs
' (middle texture) that entirely block the view.
Private Sub R_ClipSolidWallSegment(ByVal first As Long, ByVal last As Long)
    Dim nxt As Long, start As Long, cond As Boolean

    ' Find the first range that touches the range (adjacent pixels are touching).
    start = 0
    Do While ss_last(start) < first - 1
        start = start + 1
    Loop

    If first < ss_first(start) Then
        If last < ss_first(start) - 1 Then
            ' Post is entirely visible (above start), so insert a new clippost.
            R_StoreWallRange first, last
            nxt = newend
            newend = newend + 1
            Do While nxt <> start
                ss_first(nxt) = ss_first(nxt - 1)
                ss_last(nxt) = ss_last(nxt - 1)
                nxt = nxt - 1
            Loop
            ss_first(nxt) = first
            ss_last(nxt) = last
            Exit Sub
        End If

        ' There is a fragment above *start.
        R_StoreWallRange first, ss_first(start) - 1
        ' Now adjust the clip size.
        ss_first(start) = first
    End If

    ' Bottom contained in start?
    If last <= ss_last(start) Then Exit Sub

    nxt = start
    Do While last >= ss_first(nxt + 1) - 1
        ' There is a fragment between two posts.
        R_StoreWallRange ss_last(nxt) + 1, ss_first(nxt + 1) - 1
        nxt = nxt + 1
        If last <= ss_last(nxt) Then
            ' Bottom is contained in next.  Adjust the clip size.
            ss_last(start) = ss_last(nxt)
            GoTo crunch
        End If
    Loop

    ' There is a fragment after *next.
    R_StoreWallRange ss_last(nxt) + 1, last
    ' Adjust the clip size.
    ss_last(start) = last

    ' Remove start+1 to next from the clip list, because start now covers their area.
crunch:
    If nxt = start Then Exit Sub                 ' Post just extended past the bottom of one post.

    ' C: while (next++ != newend) *++start = *next;
    Do
        cond = (nxt <> newend)
        nxt = nxt + 1
        If Not cond Then Exit Do
        start = start + 1
        ss_first(start) = ss_first(nxt)
        ss_last(start) = ss_last(nxt)
    Loop
    newend = start + 1
End Sub

' R_ClipPassWallSegment - clips the given range of columns, but does not
' include it in the clip list.  Does handle windows, e.g. LineDefs with upper
' and lower texture.
Private Sub R_ClipPassWallSegment(ByVal first As Long, ByVal last As Long)
    Dim start As Long

    start = 0
    Do While ss_last(start) < first - 1
        start = start + 1
    Loop

    If first < ss_first(start) Then
        If last < ss_first(start) - 1 Then
            ' Post is entirely visible (above start).
            R_StoreWallRange first, last
            Exit Sub
        End If
        ' There is a fragment above *start.
        R_StoreWallRange first, ss_first(start) - 1
    End If

    ' Bottom contained in start?
    If last <= ss_last(start) Then Exit Sub

    Do While last >= ss_first(start + 1) - 1
        ' There is a fragment between two posts.
        R_StoreWallRange ss_last(start) + 1, ss_first(start + 1) - 1
        start = start + 1
        If last <= ss_last(start) Then Exit Sub
    Loop

    ' There is a fragment after *next.
    R_StoreWallRange ss_last(start) + 1, last
End Sub

Public Sub R_ClearClipSegs()
    ss_first(0) = -&H7FFFFFFF
    ss_last(0) = -1
    ss_first(1) = viewwidth
    ss_last(1) = &H7FFFFFFF
    newend = 2
End Sub

' R_AddLine - clips the given segment and adds any visible pieces to the line list.
Private Sub R_AddLine(ByVal line As Long)
    Dim x1 As Long, x2 As Long
    Dim angle1 As Long, angle2 As Long, span As Long, tspan As Long, clip2 As Long

    curline = line

    ' OPTIMIZE: quickly reject orthogonal back sides.
    angle1 = R_PointToAngle(vertexes(segs(line).v1).x, vertexes(segs(line).v1).y)
    angle2 = R_PointToAngle(vertexes(segs(line).v2).x, vertexes(segs(line).v2).y)

    ' Clip to view edges.
    ' OPTIMIZE: make constant out of 2*clipangle (FIELDOFVIEW).
    span = angle1 - angle2

    ' Back side? I.e. backface culling?
    If UGE(span, ANG180) Then Exit Sub

    ' Global angle needed by segcalc.
    rw_angle1 = angle1
    angle1 = angle1 - viewangle
    angle2 = angle2 - viewangle

    clip2 = 2 * clipangle
    tspan = angle1 + clipangle
    If UGT(tspan, clip2) Then
        tspan = tspan - clip2
        ' Totally off the left edge?
        If UGE(tspan, span) Then Exit Sub
        angle1 = clipangle
    End If
    tspan = clipangle - angle2
    If UGT(tspan, clip2) Then
        tspan = tspan - clip2
        ' Totally off the left edge?
        If UGE(tspan, span) Then Exit Sub
        angle2 = -clipangle
    End If

    ' The seg is in the view range, but not necessarily visible.
    angle1 = FINEANGLE(angle1 + ANG90)
    angle2 = FINEANGLE(angle2 + ANG90)
    x1 = viewangletox(angle1)
    x2 = viewangletox(angle2)

    ' Does not cross a pixel?
    If x1 = x2 Then Exit Sub

    backsector = segs(line).backsector

    ' Single sided line?
    If backsector = -1 Then GoTo clipsolid

    ' Closed door.
    If sectors(backsector).ceilingheight <= sectors(frontsector).floorheight Or _
       sectors(backsector).floorheight >= sectors(frontsector).ceilingheight Then GoTo clipsolid

    ' Window.
    If sectors(backsector).ceilingheight <> sectors(frontsector).ceilingheight Or _
       sectors(backsector).floorheight <> sectors(frontsector).floorheight Then GoTo clippass

    ' Reject empty lines used for triggers and special events.  Identical floor
    ' and ceiling on both sides, identical light levels on both sides, and no
    ' middle texture.
    If sectors(backsector).ceilingpic = sectors(frontsector).ceilingpic And _
       sectors(backsector).floorpic = sectors(frontsector).floorpic And _
       sectors(backsector).lightlevel = sectors(frontsector).lightlevel And _
       sides(segs(line).sidedef).midtexture = 0 Then
        Exit Sub
    End If

clippass:
    R_ClipPassWallSegment x1, x2 - 1
    Exit Sub

clipsolid:
    R_ClipSolidWallSegment x1, x2 - 1
End Sub

' R_CheckBBox - checks BSP node/subtree bounding box.  Returns true if some
' part of the bbox might be visible.
Private Function R_CheckBBox(ByVal node As Long, ByVal side As Long) As Boolean
    Dim boxx As Long, boxy As Long, boxpos As Long
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim angle1 As Long, angle2 As Long, span As Long, tspan As Long, clip2 As Long
    Dim start As Long, sx1 As Long, sx2 As Long

    If Not checkcoordInit Then InitCheckCoord

    ' Find the corners of the box that define the edges from current viewpoint.
    If viewx <= nodes(node).bbox(side, BOXLEFT) Then
        boxx = 0
    ElseIf viewx < nodes(node).bbox(side, BOXRIGHT) Then
        boxx = 1
    Else
        boxx = 2
    End If

    If viewy >= nodes(node).bbox(side, BOXTOP) Then
        boxy = 0
    ElseIf viewy > nodes(node).bbox(side, BOXBOTTOM) Then
        boxy = 1
    Else
        boxy = 2
    End If

    boxpos = boxy * 4 + boxx
    If boxpos = 5 Then
        R_CheckBBox = True
        Exit Function
    End If

    x1 = nodes(node).bbox(side, checkcoord(boxpos, 0))
    y1 = nodes(node).bbox(side, checkcoord(boxpos, 1))
    x2 = nodes(node).bbox(side, checkcoord(boxpos, 2))
    y2 = nodes(node).bbox(side, checkcoord(boxpos, 3))

    ' check clip list for an open space
    angle1 = R_PointToAngle(x1, y1) - viewangle
    angle2 = R_PointToAngle(x2, y2) - viewangle

    span = angle1 - angle2

    ' Sitting on a line?
    If UGE(span, ANG180) Then
        R_CheckBBox = True
        Exit Function
    End If

    clip2 = 2 * clipangle
    tspan = angle1 + clipangle
    If UGT(tspan, clip2) Then
        tspan = tspan - clip2
        ' Totally off the left edge?
        If UGE(tspan, span) Then
            R_CheckBBox = False
            Exit Function
        End If
        angle1 = clipangle
    End If
    tspan = clipangle - angle2
    If UGT(tspan, clip2) Then
        tspan = tspan - clip2
        ' Totally off the left edge?
        If UGE(tspan, span) Then
            R_CheckBBox = False
            Exit Function
        End If
        angle2 = -clipangle
    End If

    ' Find the first clippost that touches the source post (adjacent pixels
    ' are touching).
    angle1 = FINEANGLE(angle1 + ANG90)
    angle2 = FINEANGLE(angle2 + ANG90)
    sx1 = viewangletox(angle1)
    sx2 = viewangletox(angle2)

    ' Does not cross a pixel.
    If sx1 = sx2 Then
        R_CheckBBox = False
        Exit Function
    End If
    sx2 = sx2 - 1

    start = 0
    Do While ss_last(start) < sx2
        start = start + 1
    Loop

    If sx1 >= ss_first(start) And sx2 <= ss_last(start) Then
        ' The clippost contains the new span.
        R_CheckBBox = False
        Exit Function
    End If

    R_CheckBBox = True
End Function

Private Sub InitCheckCoord()
    Dim v As Variant, i As Long, j As Long
    v = Array(3, 0, 2, 1, _
              3, 0, 2, 0, _
              3, 1, 2, 0, _
              0, 0, 0, 0, _
              2, 0, 2, 1, _
              0, 0, 0, 0, _
              3, 1, 3, 0, _
              0, 0, 0, 0, _
              2, 0, 3, 1, _
              2, 1, 3, 1, _
              2, 1, 3, 0, _
              0, 0, 0, 0)
    For i = 0 To 11
        For j = 0 To 3
            checkcoord(i, j) = v(i * 4 + j)
        Next
    Next
    checkcoordInit = True
End Sub

' R_Subsector - determine floor/ceiling planes, add sprites of things in
' sector, draw one or more line segments.
Private Sub R_Subsector(ByVal num As Long)
    Dim count As Long, line As Long

    If num >= numsubsectors Then
        I_Error "R_Subsector: ss " & num & " with numss = " & numsubsectors
    End If

    sscount = sscount + 1
    frontsector = subsectors(num).sector
    count = subsectors(num).numlines
    line = subsectors(num).firstline

    If sectors(frontsector).floorheight < viewz Then
        floorplane = R_FindPlane(sectors(frontsector).floorheight, _
                                 sectors(frontsector).floorpic, _
                                 sectors(frontsector).lightlevel)
    Else
        floorplane = -1
    End If

    If sectors(frontsector).ceilingheight > viewz Or _
       sectors(frontsector).ceilingpic = skyflatnum Then
        ceilingplane = R_FindPlane(sectors(frontsector).ceilingheight, _
                                   sectors(frontsector).ceilingpic, _
                                   sectors(frontsector).lightlevel)
    Else
        ceilingplane = -1
    End If

    R_AddSprites frontsector

    Do While count > 0
        R_AddLine line
        line = line + 1
        count = count - 1
    Loop
End Sub

' RenderBSPNode - renders all subsectors below a given node, traversing
' subtree recursively.  Just call with BSP root.
Public Sub R_RenderBSPNode(ByVal bspnum As Long)
    Dim side As Long

    ' Found a subsector?
    If (bspnum And NF_SUBSECTOR) <> 0 Then
        If bspnum = -1 Then
            R_Subsector 0
        Else
            R_Subsector bspnum And &H7FFF&
        End If
        Exit Sub
    End If

    ' Decide which side the view point is on.
    side = R_PointOnSide(viewx, viewy, bspnum)

    ' Recursively divide front space.
    R_RenderBSPNode nodes(bspnum).children(side)

    ' Possibly divide back space.
    If R_CheckBBox(bspnum, side Xor 1) Then
        R_RenderBSPNode nodes(bspnum).children(side Xor 1)
    End If
End Sub
