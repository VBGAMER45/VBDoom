Attribute VB_Name = "p_sight"
' p_sight.bas - LineOfSight/Visibility checks, uses REJECT Lookup Table
' (p_sight.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' P_CheckSight
Private sightzstart As Long                      ' eye z of looker
' topslope / bottomslope (slopes to top and bottom of target) are in doomdef

Private strace As divline_t                      ' from t1 to t2
Private t2x As Long
Private t2y As Long

Private sightcounts(0 To 1) As Long

' P_DivlineSide - returns side 0 (front), 1 (back), or 2 (on).
Private Function P_DivlineSide(ByVal x As Long, ByVal y As Long, ByVal nx As Long, ByVal ny As Long, _
                               ByVal ndx As Long, ByVal ndy As Long) As Long
    Dim dx As Long, dy As Long, lft As Long, rgt As Long

    If ndx = 0 Then
        If x = nx Then
            P_DivlineSide = 2
        ElseIf x <= nx Then
            If ndy > 0 Then P_DivlineSide = 1 Else P_DivlineSide = 0
        Else
            If ndy < 0 Then P_DivlineSide = 1 Else P_DivlineSide = 0
        End If
        Exit Function
    End If

    If ndy = 0 Then
        If x = ny Then                           ' (sic: compares x with y, as in the C code)
            P_DivlineSide = 2
        ElseIf y <= ny Then
            If ndx < 0 Then P_DivlineSide = 1 Else P_DivlineSide = 0
        Else
            If ndx > 0 Then P_DivlineSide = 1 Else P_DivlineSide = 0
        End If
        Exit Function
    End If

    dx = x - nx
    dy = y - ny

    lft = ((ndy And &HFFFF0000) \ &H10000) * ((dx And &HFFFF0000) \ &H10000)
    rgt = ((dy And &HFFFF0000) \ &H10000) * ((ndx And &HFFFF0000) \ &H10000)

    If rgt < lft Then
        P_DivlineSide = 0                        ' front side
    ElseIf lft = rgt Then
        P_DivlineSide = 2
    Else
        P_DivlineSide = 1                        ' back side
    End If
End Function

' P_InterceptVector2 - returns the fractional intercept point along the
' first divline.  This is only called by the addthings and addlines
' traversers.
Private Function P_InterceptVector2(v2 As divline_t, v1 As divline_t) As Long
    Dim num As Long, den As Long

    den = FixedMul(SAR(v1.dy, 8), v2.dx) - FixedMul(SAR(v1.dx, 8), v2.dy)

    If den = 0 Then
        P_InterceptVector2 = 0
        Exit Function
    End If

    num = FixedMul(SAR(v1.x - v2.x, 8), v1.dy) + FixedMul(SAR(v2.y - v1.y, 8), v1.dx)
    P_InterceptVector2 = FixedDiv(num, den)
End Function

' P_CrossSubsector - returns true if strace crosses the given subsector
' successfully.
Private Function P_CrossSubsector(ByVal num As Long) As Boolean
    Dim seg As Long, line As Long, s1 As Long, s2 As Long, count As Long
    Dim front As Long, back As Long, otop As Long, obottom As Long
    Dim divl As divline_t, v1 As Long, v2 As Long, frac As Long, slope As Long

    If num >= numsubsectors Then
        I_Error "P_CrossSubsector: ss " & num & " with numss = " & numsubsectors
    End If

    ' check lines
    count = subsectors(num).numlines
    seg = subsectors(num).firstline

    Do While count > 0
        line = segs(seg).linedef

        ' allready checked other side?
        If lines(line).validcount = validcount Then GoTo nextseg

        lines(line).validcount = validcount

        v1 = lines(line).v1
        v2 = lines(line).v2
        s1 = P_DivlineSide(vertexes(v1).x, vertexes(v1).y, strace.x, strace.y, strace.dx, strace.dy)
        s2 = P_DivlineSide(vertexes(v2).x, vertexes(v2).y, strace.x, strace.y, strace.dx, strace.dy)

        ' line isn't crossed?
        If s1 = s2 Then GoTo nextseg

        divl.x = vertexes(v1).x
        divl.y = vertexes(v1).y
        divl.dx = vertexes(v2).x - vertexes(v1).x
        divl.dy = vertexes(v2).y - vertexes(v1).y
        s1 = P_DivlineSide(strace.x, strace.y, divl.x, divl.y, divl.dx, divl.dy)
        s2 = P_DivlineSide(t2x, t2y, divl.x, divl.y, divl.dx, divl.dy)

        ' line isn't crossed?
        If s1 = s2 Then GoTo nextseg

        ' stop because it is not two sided anyway
        ' might do this after updating validcount?
        If (lines(line).flags And ML_TWOSIDED) = 0 Then
            P_CrossSubsector = False
            Exit Function
        End If

        ' crosses a two sided line
        front = segs(seg).frontsector
        back = segs(seg).backsector

        ' no wall to block sight with?
        If sectors(front).floorheight = sectors(back).floorheight And _
           sectors(front).ceilingheight = sectors(back).ceilingheight Then GoTo nextseg

        ' possible occluder
        ' because of ceiling height differences
        If sectors(front).ceilingheight < sectors(back).ceilingheight Then
            otop = sectors(front).ceilingheight
        Else
            otop = sectors(back).ceilingheight
        End If

        ' because of ceiling height differences
        If sectors(front).floorheight > sectors(back).floorheight Then
            obottom = sectors(front).floorheight
        Else
            obottom = sectors(back).floorheight
        End If

        ' quick test for totally closed doors
        If obottom >= otop Then
            P_CrossSubsector = False             ' stop
            Exit Function
        End If

        frac = P_InterceptVector2(strace, divl)

        If sectors(front).floorheight <> sectors(back).floorheight Then
            slope = FixedDiv(obottom - sightzstart, frac)
            If slope > bottomslope Then bottomslope = slope
        End If

        If sectors(front).ceilingheight <> sectors(back).ceilingheight Then
            slope = FixedDiv(otop - sightzstart, frac)
            If slope < topslope Then topslope = slope
        End If

        If topslope <= bottomslope Then
            P_CrossSubsector = False             ' stop
            Exit Function
        End If
nextseg:
        seg = seg + 1
        count = count - 1
    Loop
    ' passed the subsector ok
    P_CrossSubsector = True
End Function

' P_CrossBSPNode - returns true if strace crosses the given node successfully.
Private Function P_CrossBSPNode(ByVal bspnum As Long) As Boolean
    Dim side As Long

    If (bspnum And NF_SUBSECTOR) <> 0 Then
        If bspnum = -1 Then
            P_CrossBSPNode = P_CrossSubsector(0)
        Else
            P_CrossBSPNode = P_CrossSubsector(bspnum And &H7FFF&)
        End If
        Exit Function
    End If

    ' decide which side the start point is on
    side = P_DivlineSide(strace.x, strace.y, nodes(bspnum).x, nodes(bspnum).y, nodes(bspnum).dx, nodes(bspnum).dy)
    If side = 2 Then side = 0                    ' an "on" should cross both sides

    ' cross the starting side
    If Not P_CrossBSPNode(nodes(bspnum).children(side)) Then
        P_CrossBSPNode = False
        Exit Function
    End If

    ' the partition plane is crossed here
    If side = P_DivlineSide(t2x, t2y, nodes(bspnum).x, nodes(bspnum).y, nodes(bspnum).dx, nodes(bspnum).dy) Then
        ' the line doesn't touch the other side
        P_CrossBSPNode = True
        Exit Function
    End If

    ' cross the ending side
    P_CrossBSPNode = P_CrossBSPNode(nodes(bspnum).children(side Xor 1))
End Function

' P_CheckSight - returns true if a straight line between t1 and t2 is
' unobstructed.  Uses REJECT.
Public Function P_CheckSight(ByVal t1 As Long, ByVal t2 As Long) As Boolean
    Dim s1 As Long, s2 As Long, pnum As Long, bytenum As Long, bitnum As Long

    ' First check for trivial rejection.

    ' Determine subsector entries in REJECT table.
    s1 = subsectors(mobjs(t1).subsector).sector
    s2 = subsectors(mobjs(t2).subsector).sector
    pnum = s1 * numsectors + s2
    bytenum = pnum \ 8
    bitnum = pow2(pnum And 7)

    ' Check in REJECT table.
    If (rejectmatrix(bytenum) And bitnum) <> 0 Then
        sightcounts(0) = sightcounts(0) + 1

        ' can't possibly be connected
        P_CheckSight = False
        Exit Function
    End If

    ' An unobstructed LOS is possible.
    ' Now look from eyes of t1 to any part of t2.
    sightcounts(1) = sightcounts(1) + 1

    validcount = validcount + 1

    sightzstart = mobjs(t1).z + mobjs(t1).height - SAR(mobjs(t1).height, 2)
    topslope = (mobjs(t2).z + mobjs(t2).height) - sightzstart
    bottomslope = mobjs(t2).z - sightzstart

    strace.x = mobjs(t1).x
    strace.y = mobjs(t1).y
    t2x = mobjs(t2).x
    t2y = mobjs(t2).y
    strace.dx = mobjs(t2).x - mobjs(t1).x
    strace.dy = mobjs(t2).y - mobjs(t1).y

    ' the head node is the last node output
    P_CheckSight = P_CrossBSPNode(numnodes - 1)
End Function
