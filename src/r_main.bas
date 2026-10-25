Attribute VB_Name = "r_main"
' r_main.bas - rendering main loop and setup functions, utility functions
' (BSP, geometry, trigonometry) (r_main.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public Const FIELDOFVIEW As Long = 2048          ' fineangles in the SCREENWIDTH wide window
Public Const LIGHTLEVELS As Long = 16
Public Const LIGHTSEGSHIFT As Long = 4
Public Const MAXLIGHTSCALE As Long = 48
Public Const LIGHTSCALESHIFT As Long = 12
Public Const MAXLIGHTZ As Long = 128
Public Const LIGHTZSHIFT As Long = 20
Public Const NUMCOLORMAPS As Long = 32
Private Const DISTMAP As Long = 2

Public fixedcolormap As Long                     ' offset in colormaps(), 0 = none

Public centerx As Long
Public centery As Long
Public centerxfrac As Long
Public centeryfrac As Long
Public projection As Long

Public framecount As Long                        ' just for profiling purposes
Public sscount As Long

Public viewcos As Long
Public viewsin As Long

' The xtoviewangle() table maps a screen pixel to the lowest viewangle that
' maps back to x ranges from clipangle to -clipangle.
Public clipangle As Long

' The viewangletox() table maps a FINEANGLE to the x coordinate.  (Sized for
' the whole fine circle so a stray index cannot run off the end.)
Public viewangletox(0 To 8191) As Long
Public xtoviewangle(0 To 320) As Long

' lighting: values are offsets into colormaps()
Public scalelight(0 To 15, 0 To 47) As Long      ' [LIGHTLEVELS][MAXLIGHTSCALE]
Public zlight(0 To 15, 0 To 127) As Long         ' [LIGHTLEVELS][MAXLIGHTZ]
Public walllights As Long                        ' scalelight row; -1 = fixedcolormap

Private setblocks As Long
Private setdetail As Long

' R_PointOnSide - traverse BSP (sub) tree, check point against partition plane.
' Returns side 0 (front) or 1 (back).
Public Function R_PointOnSide(ByVal x As Long, ByVal y As Long, ByVal node As Long) As Long
    Dim dx As Long, dy As Long, lft As Long, rgt As Long
    Dim ndx As Long, ndy As Long

    ndx = nodes(node).dx
    ndy = nodes(node).dy
    If ndx = 0 Then
        If x <= nodes(node).x Then
            If ndy > 0 Then R_PointOnSide = 1 Else R_PointOnSide = 0
        Else
            If ndy < 0 Then R_PointOnSide = 1 Else R_PointOnSide = 0
        End If
        Exit Function
    End If
    If ndy = 0 Then
        If y <= nodes(node).y Then
            If ndx < 0 Then R_PointOnSide = 1 Else R_PointOnSide = 0
        Else
            If ndx > 0 Then R_PointOnSide = 1 Else R_PointOnSide = 0
        End If
        Exit Function
    End If

    dx = x - nodes(node).x
    dy = y - nodes(node).y

    ' Try to quickly decide by looking at sign bits.
    If (ndy Xor ndx Xor dx Xor dy) < 0 Then
        If (ndy Xor dx) < 0 Then
            R_PointOnSide = 1                    ' (left is negative)
        Else
            R_PointOnSide = 0
        End If
        Exit Function
    End If

    lft = FixedMul((ndy And &HFFFF0000) \ &H10000, dx)
    rgt = FixedMul(dy, (ndx And &HFFFF0000) \ &H10000)
    If rgt < lft Then
        R_PointOnSide = 0                        ' front side
    Else
        R_PointOnSide = 1                        ' back side
    End If
End Function

Public Function R_PointOnSegSide(ByVal x As Long, ByVal y As Long, ByVal seg As Long) As Long
    Dim lx As Long, ly As Long, ldx As Long, ldy As Long
    Dim dx As Long, dy As Long, lft As Long, rgt As Long

    lx = vertexes(segs(seg).v1).x
    ly = vertexes(segs(seg).v1).y
    ldx = vertexes(segs(seg).v2).x - lx
    ldy = vertexes(segs(seg).v2).y - ly

    If ldx = 0 Then
        If x <= lx Then
            If ldy > 0 Then R_PointOnSegSide = 1 Else R_PointOnSegSide = 0
        Else
            If ldy < 0 Then R_PointOnSegSide = 1 Else R_PointOnSegSide = 0
        End If
        Exit Function
    End If
    If ldy = 0 Then
        If y <= ly Then
            If ldx < 0 Then R_PointOnSegSide = 1 Else R_PointOnSegSide = 0
        Else
            If ldx > 0 Then R_PointOnSegSide = 1 Else R_PointOnSegSide = 0
        End If
        Exit Function
    End If

    dx = x - lx
    dy = y - ly

    If (ldy Xor ldx Xor dx Xor dy) < 0 Then
        If (ldy Xor dx) < 0 Then
            R_PointOnSegSide = 1
        Else
            R_PointOnSegSide = 0
        End If
        Exit Function
    End If

    lft = FixedMul((ldy And &HFFFF0000) \ &H10000, dx)
    rgt = FixedMul(dy, (ldx And &HFFFF0000) \ &H10000)
    If rgt < lft Then
        R_PointOnSegSide = 0
    Else
        R_PointOnSegSide = 1
    End If
End Function

' R_PointToAngle - to get a global angle from cartesian coordinates, the
' coordinates are flipped until they are in the first octant of the
' coordinate system, then the y (<=x) is scaled and divided by x to get a
' tangent (slope) value which is looked up in the tantoangle() table.
Public Function R_PointToAngle(ByVal x As Long, ByVal y As Long) As Long
    x = x - viewx
    y = y - viewy

    If x = 0 And y = 0 Then
        R_PointToAngle = 0
        Exit Function
    End If

    If x >= 0 Then
        ' x >=0
        If y >= 0 Then
            ' y>= 0
            If x > y Then
                R_PointToAngle = tantoangle(SlopeDiv(y, x))                  ' octant 0
            Else
                R_PointToAngle = ANG90 - 1 - tantoangle(SlopeDiv(x, y))      ' octant 1
            End If
        Else
            ' y<0
            y = -y
            If x > y Then
                R_PointToAngle = -tantoangle(SlopeDiv(y, x))                 ' octant 8
            Else
                R_PointToAngle = ANG270 + tantoangle(SlopeDiv(x, y))         ' octant 7
            End If
        End If
    Else
        ' x<0
        x = -x
        If y >= 0 Then
            ' y>= 0
            If x > y Then
                R_PointToAngle = ANG180 - 1 - tantoangle(SlopeDiv(y, x))     ' octant 3
            Else
                R_PointToAngle = ANG90 + tantoangle(SlopeDiv(x, y))          ' octant 2
            End If
        Else
            ' y<0
            y = -y
            If x > y Then
                R_PointToAngle = ANG180 + tantoangle(SlopeDiv(y, x))         ' octant 4
            Else
                R_PointToAngle = ANG270 - 1 - tantoangle(SlopeDiv(x, y))     ' octant 5
            End If
        End If
    End If
End Function

Public Function R_PointToAngle2(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
    viewx = x1
    viewy = y1
    R_PointToAngle2 = R_PointToAngle(x2, y2)
End Function

Public Function R_PointToDist(ByVal x As Long, ByVal y As Long) As Long
    Dim angle As Long, dx As Long, dy As Long, temp As Long

    dx = Abs(x - viewx)
    dy = Abs(y - viewy)
    If dy > dx Then
        temp = dx
        dx = dy
        dy = temp
    End If
    If dx = 0 Then
        R_PointToDist = 0
        Exit Function
    End If
    angle = FINEANGLE(tantoangle(FixedDiv(dy, dx) \ 32) + ANG90)
    ' use as cosine
    R_PointToDist = FixedDiv(dx, finesine(angle))
End Function

' R_ScaleFromGlobalAngle - returns the texture mapping scale for the current
' line (horizontal span) at the given angle.  rw_distance must be calculated first.
Public Function R_ScaleFromGlobalAngle(ByVal visangle As Long) As Long
    Dim scl As Long, anglea As Long, angleb As Long
    Dim sinea As Long, sineb As Long, num As Long, den As Long

    anglea = ANG90 + (visangle - viewangle)
    angleb = ANG90 + (visangle - rw_normalangle)

    ' both sines are allways positive
    sinea = finesine(FINEANGLE(anglea))
    sineb = finesine(FINEANGLE(angleb))
    num = FixedMul(projection, sineb) * pow2(detailshift)
    den = FixedMul(rw_distance, sinea)

    If den > ((num And &HFFFF0000) \ &H10000) Then
        scl = FixedDiv(num, den)
        If scl > 64 * FRACUNIT Then
            scl = 64 * FRACUNIT
        ElseIf scl < 256 Then
            scl = 256
        End If
    Else
        scl = 64 * FRACUNIT
    End If
    R_ScaleFromGlobalAngle = scl
End Function

Private Sub R_InitTextureMapping()
    Dim i As Long, x As Long, t As Long, focallength As Long

    ' Use tangent table to generate viewangletox:
    '  viewangletox will give the next greatest x
    '  after the view angle.
    '
    ' Calc focallength
    '  so FIELDOFVIEW angles covers SCREENWIDTH.
    focallength = FixedDiv(centerxfrac, finetangent(FINEANGLES \ 4 + FIELDOFVIEW \ 2))

    For i = 0 To FINEANGLES \ 2 - 1
        If finetangent(i) > FRACUNIT * 2 Then
            t = -1
        ElseIf finetangent(i) < -FRACUNIT * 2 Then
            t = viewwidth + 1
        Else
            t = FixedMul(finetangent(i), focallength)
            t = (centerxfrac - t + FRACUNIT - 1)
            t = (t And &HFFFF0000) \ &H10000
            If t < -1 Then
                t = -1
            ElseIf t > viewwidth + 1 Then
                t = viewwidth + 1
            End If
        End If
        viewangletox(i) = t
    Next

    ' Scan viewangletox[] to generate xtoviewangle[]:
    '  xtoviewangle will give the smallest view angle
    '  that maps to x.
    For x = 0 To viewwidth
        i = 0
        Do While viewangletox(i) > x
            i = i + 1
        Loop
        xtoviewangle(x) = i * pow2(ANGLETOFINESHIFT) - ANG90
    Next

    ' Take out the fencepost cases from viewangletox.
    For i = 0 To FINEANGLES \ 2 - 1
        If viewangletox(i) = -1 Then
            viewangletox(i) = 0
        ElseIf viewangletox(i) = viewwidth + 1 Then
            viewangletox(i) = viewwidth
        End If
    Next

    clipangle = xtoviewangle(0)
End Sub

' R_InitLightTables - only inits the zlight table, because the scalelight
' table changes with view size.
Private Sub R_InitLightTables()
    Dim i As Long, j As Long, level As Long, startmap As Long, scl As Long

    ' Calculate the light levels to use for each level / distance combination.
    For i = 0 To LIGHTLEVELS - 1
        startmap = ((LIGHTLEVELS - 1 - i) * 2) * NUMCOLORMAPS \ LIGHTLEVELS
        For j = 0 To MAXLIGHTZ - 1
            scl = FixedDiv((SCREENWIDTH \ 2 * FRACUNIT), (j + 1) * pow2(LIGHTZSHIFT))
            scl = SAR(scl, LIGHTSCALESHIFT)
            level = startmap - scl \ DISTMAP
            If level < 0 Then level = 0
            If level >= NUMCOLORMAPS Then level = NUMCOLORMAPS - 1
            zlight(i, j) = level * 256
        Next
    Next
End Sub

' R_SetViewSize - do not really change anything here, because it might be in
' the middle of a refresh.  The change will take effect next refresh.
Public Sub R_SetViewSize(ByVal blocks As Long, ByVal detail As Long)
    setsizeneeded = True
    setblocks = blocks
    setdetail = detail
End Sub

Public Sub R_ExecuteSetViewSize()
    Dim cosadj As Long, dy As Long, i As Long, j As Long, level As Long, startmap As Long

    setsizeneeded = False

    If setblocks = 11 Then
        scaledviewwidth = SCREENWIDTH
        viewheight = SCREENHEIGHT
    Else
        scaledviewwidth = setblocks * 32
        viewheight = ((setblocks * 168) \ 10) And (Not 7)
    End If

    detailshift = setdetail
    viewwidth = scaledviewwidth \ pow2(detailshift)

    centery = viewheight \ 2
    centerx = viewwidth \ 2
    centerxfrac = centerx * FRACUNIT
    centeryfrac = centery * FRACUNIT
    projection = centerxfrac

    colfunc = COLF_BASE

    R_InitBuffer scaledviewwidth, viewheight

    R_InitTextureMapping

    ' psprite scales
    pspritescale = FRACUNIT * viewwidth \ SCREENWIDTH
    pspriteiscale = FRACUNIT * SCREENWIDTH \ viewwidth

    ' thing clipping
    For i = 0 To viewwidth - 1
        openings(SCREENHEIGHTARRAY + i) = viewheight
    Next

    ' planes
    For i = 0 To viewheight - 1
        dy = ((i - viewheight \ 2) * FRACUNIT) + FRACUNIT \ 2
        dy = Abs(dy)
        yslope(i) = FixedDiv((viewwidth * pow2(detailshift)) \ 2 * FRACUNIT, dy)
    Next

    For i = 0 To viewwidth - 1
        cosadj = Abs(finecosine(FINEANGLE(xtoviewangle(i))))
        distscale(i) = FixedDiv(FRACUNIT, cosadj)
    Next

    ' Calculate the light levels to use for each level / scale combination.
    For i = 0 To LIGHTLEVELS - 1
        startmap = ((LIGHTLEVELS - 1 - i) * 2) * NUMCOLORMAPS \ LIGHTLEVELS
        For j = 0 To MAXLIGHTSCALE - 1
            level = startmap - j * SCREENWIDTH \ (viewwidth * pow2(detailshift)) \ DISTMAP
            If level < 0 Then level = 0
            If level >= NUMCOLORMAPS Then level = NUMCOLORMAPS - 1
            scalelight(i, j) = level * 256
        Next
    Next
End Sub

Public Sub R_Init()
    R_InitData
    I_Printf "  R_InitData"
    R_SetViewSize screenblocks, detailLevel
    R_InitPlanes
    I_Printf "  R_InitPlanes"
    R_InitLightTables
    I_Printf "  R_InitLightTables"
    R_InitSkyMap
    I_Printf "  R_InitSkyMap"
    R_InitTranslationTables
    I_Printf "  R_InitTranslationsTables"
    framecount = 0
End Sub

' R_PointInSubsector
Public Function R_PointInSubsector(ByVal x As Long, ByVal y As Long) As Long
    Dim side As Long, nodenum As Long

    ' single subsector is a special case
    If numnodes = 0 Then
        R_PointInSubsector = 0
        Exit Function
    End If

    nodenum = numnodes - 1
    Do While (nodenum And NF_SUBSECTOR) = 0
        side = R_PointOnSide(x, y, nodenum)
        nodenum = nodes(nodenum).children(side)
    Loop
    R_PointInSubsector = nodenum And &H7FFF&
End Function

Private Sub R_SetupFrame(ByVal player As Long)
    Dim mo As Long
    viewplayer = player
    mo = players(player).mo
    viewx = mobjs(mo).x
    viewy = mobjs(mo).y
    viewangle = mobjs(mo).angle + viewangleoffset
    extralight = players(player).extralight

    viewz = players(player).viewz

    viewsin = finesine(FINEANGLE(viewangle))
    viewcos = finecosine(FINEANGLE(viewangle))

    sscount = 0

    If players(player).fixedcolormap <> 0 Then
        fixedcolormap = players(player).fixedcolormap * 256
        walllights = -1                          ' C: walllights = scalelightfixed
    Else
        fixedcolormap = 0
    End If

    framecount = framecount + 1
    validcount = validcount + 1
End Sub

Public Sub R_RenderPlayerView(ByVal playernum As Long)
    R_SetupFrame playernum

    ' Clear buffers.
    R_ClearClipSegs
    R_ClearDrawSegs
    R_ClearPlanes
    R_ClearSprites

    ' check for new console commands.
    NetUpdate

    ' The head node is the last node output.
    dbgPhase = "R_RenderBSPNode"
    R_RenderBSPNode numnodes - 1

    ' Check for new console commands.
    NetUpdate

    dbgPhase = "R_DrawPlanes"
    R_DrawPlanes

    ' Check for new console commands.
    NetUpdate

    dbgPhase = "R_DrawMasked"
    R_DrawMasked
    dbgPhase = "R_RenderPlayerView done"

    ' Check for new console commands.
    NetUpdate
End Sub
