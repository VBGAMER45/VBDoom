Attribute VB_Name = "r_segs"
' r_segs.bas - all the clipping: columns, horizontal spans, sky columns (r_segs.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const HEIGHTBITS As Long = 12
Private Const HEIGHTUNIT As Long = 4096
Public Const MAXSHORT As Long = 32767

' OPTIMIZE: closed two sided lines as single sided

' True if any of the segs textures might be visible.
Public segtextured As Boolean

' False if the back side is the same plane.
Public markfloor As Boolean
Public markceiling As Boolean

Public maskedtexture As Boolean
Public toptexture As Long
Public bottomtexture As Long
Public midtexture As Long

Public rw_normalangle As Long
' angle to line origin
Public rw_angle1 As Long

' regular wall
Public rw_x As Long
Public rw_stopx As Long
Public rw_centerangle As Long
Public rw_offset As Long
Public rw_distance As Long
Public rw_scale As Long
Public rw_scalestep As Long
Public rw_midtexturemid As Long
Public rw_toptexturemid As Long
Public rw_bottomtexturemid As Long

Public worldtop As Long
Public worldbottom As Long
Public worldhigh As Long
Public worldlow As Long

Public pixhigh As Long
Public pixlow As Long
Public pixhighstep As Long
Public pixlowstep As Long

Public topfrac As Long
Public topstep As Long

Public bottomfrac As Long
Public bottomstep As Long

Public maskedtexturecol As Long                  ' openings() offset (C: short*)

' 0xffffffffu / (unsigned)scale, as fixed_t
Public Function InvScale(ByVal s As Long) As Long
    Dim d As Double
    If s = 0 Then
        InvScale = -1
        Exit Function
    End If
    d = Int(4294967295# / UDbl(s))
    If d > 2147483647# Then d = d - 4294967296#
    InvScale = d
End Function

' R_RenderMaskedSegRange
Public Sub R_RenderMaskedSegRange(ByVal ds As Long, ByVal x1 As Long, ByVal x2 As Long)
    Dim index As Long, lightnum As Long, texnum As Long, x As Long
    Dim col As Long, texcol As Long, v As Long

    ' Calculate light table.
    ' Use different light tables for horizontal / vertical / diagonal.
    ' OPTIMIZE: get rid of LIGHTSEGSHIFT globally
    curline = drawsegs(ds).curline
    frontsector = segs(curline).frontsector
    backsector = segs(curline).backsector
    texnum = texturetranslation(sides(segs(curline).sidedef).midtexture)

    lightnum = (sectors(frontsector).lightlevel \ 16) + extralight

    If vertexes(segs(curline).v1).y = vertexes(segs(curline).v2).y Then
        lightnum = lightnum - 1
    ElseIf vertexes(segs(curline).v1).x = vertexes(segs(curline).v2).x Then
        lightnum = lightnum + 1
    End If

    If lightnum < 0 Then
        walllights = 0
    ElseIf lightnum >= LIGHTLEVELS Then
        walllights = LIGHTLEVELS - 1
    Else
        walllights = lightnum
    End If

    maskedtexturecol = drawsegs(ds).maskedtexturecol

    rw_scalestep = drawsegs(ds).scalestep
    spryscale = drawsegs(ds).scale1 + (x1 - drawsegs(ds).x1) * rw_scalestep
    mfloorclip = drawsegs(ds).sprbottomclip
    mceilingclip = drawsegs(ds).sprtopclip

    ' find positioning
    If (lines(segs(curline).linedef).flags And ML_DONTPEGBOTTOM) <> 0 Then
        If sectors(frontsector).floorheight > sectors(backsector).floorheight Then
            dc_texturemid = sectors(frontsector).floorheight
        Else
            dc_texturemid = sectors(backsector).floorheight
        End If
        dc_texturemid = dc_texturemid + textureheight(texnum) - viewz
    Else
        If sectors(frontsector).ceilingheight < sectors(backsector).ceilingheight Then
            dc_texturemid = sectors(frontsector).ceilingheight
        Else
            dc_texturemid = sectors(backsector).ceilingheight
        End If
        dc_texturemid = dc_texturemid - viewz
    End If
    dc_texturemid = dc_texturemid + sides(segs(curline).sidedef).rowoffset

    If fixedcolormap <> 0 Then dc_colormap = fixedcolormap

    ' draw the columns
    For x = x1 To x2
        dc_x = x
        ' calculate lighting
        texcol = openings(maskedtexturecol + x)
        If texcol <> MAXSHORT Then
            If fixedcolormap = 0 Then
                index = SAR(spryscale, LIGHTSCALESHIFT)
                If index >= MAXLIGHTSCALE Or index < 0 Then index = MAXLIGHTSCALE - 1
                dc_colormap = scalelight(walllights, index)
            End If

            sprtopscreen = centeryfrac - FixedMul(dc_texturemid, spryscale)
            dc_iscale = InvScale(spryscale)

            ' draw the texture
            If R_IsCompositeColumn(texnum, texcol) Then
                R_DrawSolidMaskedColumn R_GetColumn(texnum, texcol), textureheight(texnum)
            Else
                col = R_GetColumn(texnum, texcol) - 3
                R_DrawMaskedColumn col
            End If
            openings(maskedtexturecol + x) = MAXSHORT
        End If
        spryscale = spryscale + rw_scalestep
    Next
End Sub

' R_RenderSegLoop - draws zero, one, or two textures (and possibly a masked
' texture) for walls.  Can draw or mark the starting pixel of floor and
' ceiling textures.  CALLED: CORE LOOPING ROUTINE.
Private Sub R_RenderSegLoop()
    Dim angle As Long, index As Long, yl As Long, yh As Long, mid As Long
    Dim texturecolumn As Long, top As Long, bottom As Long
    Dim cclip As Long, fclip As Long

    Do While rw_x < rw_stopx
        cclip = ceilingclip(rw_x)
        fclip = floorclip(rw_x)

        ' mark floor / ceiling areas
        yl = ((topfrac + HEIGHTUNIT - 1) And &HFFFFF000) \ HEIGHTUNIT

        ' no space above wall?
        If yl < cclip + 1 Then yl = cclip + 1

        If markceiling Then
            top = cclip + 1
            bottom = yl - 1
            If bottom >= fclip Then bottom = fclip - 1
            If top <= bottom Then
                vp_top(ceilingplane * VPSTRIDE + rw_x + 1) = top
                vp_bottom(ceilingplane * VPSTRIDE + rw_x + 1) = bottom
            End If
        End If

        yh = (bottomfrac And &HFFFFF000) \ HEIGHTUNIT

        If yh >= fclip Then yh = fclip - 1

        If markfloor Then
            top = yh + 1
            bottom = fclip - 1
            If top <= cclip Then top = cclip + 1
            If top <= bottom Then
                vp_top(floorplane * VPSTRIDE + rw_x + 1) = top
                vp_bottom(floorplane * VPSTRIDE + rw_x + 1) = bottom
            End If
        End If

        ' texturecolumn and lighting are independent of wall tiers
        If segtextured Then
            ' calculate texture offset
            angle = (((rw_centerangle + xtoviewangle(rw_x)) And &HFFF80000) \ &H80000) And &H1FFF&
            texturecolumn = rw_offset - FixedMul(finetangent(angle), rw_distance)
            texturecolumn = (texturecolumn And &HFFFF0000) \ &H10000
            ' calculate lighting
            index = (rw_scale And &HFFFFF000) \ 4096
            If index >= MAXLIGHTSCALE Or index < 0 Then index = MAXLIGHTSCALE - 1

            If walllights < 0 Then
                dc_colormap = fixedcolormap
            Else
                dc_colormap = scalelight(walllights, index)
            End If
            dc_x = rw_x
            dc_iscale = InvScale(rw_scale)
        End If

        ' draw the wall tiers
        If midtexture <> 0 Then
            ' single sided line
            dc_yl = yl
            dc_yh = yh
            dc_texturemid = rw_midtexturemid
            dc_source = R_GetColumn(midtexture, texturecolumn)
            R_ColFunc
            ceilingclip(rw_x) = viewheight
            floorclip(rw_x) = -1
        Else
            ' two sided line
            If toptexture <> 0 Then
                ' top wall
                mid = (pixhigh And &HFFFFF000) \ HEIGHTUNIT
                pixhigh = pixhigh + pixhighstep

                If mid >= fclip Then mid = fclip - 1

                If mid >= yl Then
                    dc_yl = yl
                    dc_yh = mid
                    dc_texturemid = rw_toptexturemid
                    dc_source = R_GetColumn(toptexture, texturecolumn)
                    R_ColFunc
                    ceilingclip(rw_x) = mid
                Else
                    ceilingclip(rw_x) = yl - 1
                End If
            Else
                ' no top wall
                If markceiling Then ceilingclip(rw_x) = yl - 1
            End If

            If bottomtexture <> 0 Then
                ' bottom wall
                mid = ((pixlow + HEIGHTUNIT - 1) And &HFFFFF000) \ HEIGHTUNIT
                pixlow = pixlow + pixlowstep

                ' no space above wall?
                If mid <= ceilingclip(rw_x) Then mid = ceilingclip(rw_x) + 1

                If mid <= yh Then
                    dc_yl = mid
                    dc_yh = yh
                    dc_texturemid = rw_bottomtexturemid
                    dc_source = R_GetColumn(bottomtexture, texturecolumn)
                    R_ColFunc
                    floorclip(rw_x) = mid
                Else
                    floorclip(rw_x) = yh + 1
                End If
            Else
                ' no bottom wall
                If markfloor Then floorclip(rw_x) = yh + 1
            End If

            If maskedtexture Then
                ' save texturecol for backdrawing of masked mid texture
                openings(maskedtexturecol + rw_x) = texturecolumn
            End If
        End If

        rw_scale = rw_scale + rw_scalestep
        topfrac = topfrac + topstep
        bottomfrac = bottomfrac + bottomstep
        rw_x = rw_x + 1
    Loop
End Sub

' R_StoreWallRange - a wall segment will be drawn between start and stop pixels (inclusive).
Public Sub R_StoreWallRange(ByVal start As Long, ByVal stopx As Long)
    Dim hyp As Long, sineval As Long, distangle As Long, offsetangle As Long
    Dim vtop As Long, lightnum As Long, fs As Long, bs As Long, ds As Long

    R_CheckDrawSegs
    ds = ds_p

    If start >= viewwidth Or start > stopx Then
        I_Error "Bad R_RenderWallRange: " & start & " to " & stopx
    End If

    sidedef = segs(curline).sidedef
    linedef = segs(curline).linedef
    fs = frontsector
    bs = backsector

    ' mark the segment as visible for auto map
    lines(linedef).flags = lines(linedef).flags Or ML_MAPPED

    ' calculate rw_distance for scale calculation
    rw_normalangle = segs(curline).angle + ANG90
    offsetangle = rw_normalangle - rw_angle1
    If offsetangle < 0 Then offsetangle = -offsetangle    ' C: abs()

    If UGT(offsetangle, ANG90) Then offsetangle = ANG90

    distangle = ANG90 - offsetangle
    hyp = R_PointToDist(vertexes(segs(curline).v1).x, vertexes(segs(curline).v1).y)
    sineval = finesine(FINEANGLE(distangle))
    rw_distance = FixedMul(hyp, sineval)

    rw_x = start
    drawsegs(ds).x1 = start
    drawsegs(ds).x2 = stopx
    drawsegs(ds).curline = curline
    rw_stopx = stopx + 1

    ' calculate scale at both ends and step
    rw_scale = R_ScaleFromGlobalAngle(viewangle + xtoviewangle(start))
    drawsegs(ds).scale1 = rw_scale

    If stopx > start Then
        drawsegs(ds).scale2 = R_ScaleFromGlobalAngle(viewangle + xtoviewangle(stopx))
        rw_scalestep = (drawsegs(ds).scale2 - rw_scale) \ (stopx - start)
        drawsegs(ds).scalestep = rw_scalestep
    Else
        drawsegs(ds).scale2 = drawsegs(ds).scale1
    End If

    ' calculate texture boundaries and decide if floor / ceiling marks are needed
    worldtop = sectors(fs).ceilingheight - viewz
    worldbottom = sectors(fs).floorheight - viewz

    midtexture = 0
    toptexture = 0
    bottomtexture = 0
    maskedtexture = False
    drawsegs(ds).maskedtexturecol = -1

    If bs = -1 Then
        ' single sided line
        midtexture = texturetranslation(sides(sidedef).midtexture)
        ' a single sided line is terminal, so it must mark ends
        markfloor = True
        markceiling = True
        If (lines(linedef).flags And ML_DONTPEGBOTTOM) <> 0 Then
            vtop = sectors(fs).floorheight + textureheight(sides(sidedef).midtexture)
            ' bottom of texture at bottom
            rw_midtexturemid = vtop - viewz
        Else
            ' top of texture at top
            rw_midtexturemid = worldtop
        End If
        rw_midtexturemid = rw_midtexturemid + sides(sidedef).rowoffset

        drawsegs(ds).silhouette = SIL_BOTH
        drawsegs(ds).sprtopclip = SCREENHEIGHTARRAY
        drawsegs(ds).sprbottomclip = NEGONEARRAY
        drawsegs(ds).bsilheight = MAXINT
        drawsegs(ds).tsilheight = MININT
    Else
        ' two sided line
        drawsegs(ds).sprtopclip = -1
        drawsegs(ds).sprbottomclip = -1
        drawsegs(ds).silhouette = 0

        If sectors(fs).floorheight > sectors(bs).floorheight Then
            drawsegs(ds).silhouette = SIL_BOTTOM
            drawsegs(ds).bsilheight = sectors(fs).floorheight
        ElseIf sectors(bs).floorheight > viewz Then
            drawsegs(ds).silhouette = SIL_BOTTOM
            drawsegs(ds).bsilheight = MAXINT
            ' ds_p->sprbottomclip = negonearray;
        End If

        If sectors(fs).ceilingheight < sectors(bs).ceilingheight Then
            drawsegs(ds).silhouette = drawsegs(ds).silhouette Or SIL_TOP
            drawsegs(ds).tsilheight = sectors(fs).ceilingheight
        ElseIf sectors(bs).ceilingheight < viewz Then
            drawsegs(ds).silhouette = drawsegs(ds).silhouette Or SIL_TOP
            drawsegs(ds).tsilheight = MININT
            ' ds_p->sprtopclip = screenheightarray;
        End If

        If sectors(bs).ceilingheight <= sectors(fs).floorheight Then
            drawsegs(ds).sprbottomclip = NEGONEARRAY
            drawsegs(ds).bsilheight = MAXINT
            drawsegs(ds).silhouette = drawsegs(ds).silhouette Or SIL_BOTTOM
        End If

        If sectors(bs).floorheight >= sectors(fs).ceilingheight Then
            drawsegs(ds).sprtopclip = SCREENHEIGHTARRAY
            drawsegs(ds).tsilheight = MININT
            drawsegs(ds).silhouette = drawsegs(ds).silhouette Or SIL_TOP
        End If

        worldhigh = sectors(bs).ceilingheight - viewz
        worldlow = sectors(bs).floorheight - viewz

        ' hack to allow height changes in outdoor areas
        If sectors(fs).ceilingpic = skyflatnum And sectors(bs).ceilingpic = skyflatnum Then
            worldtop = worldhigh
        End If

        If worldlow <> worldbottom Or _
           sectors(bs).floorpic <> sectors(fs).floorpic Or _
           sectors(bs).lightlevel <> sectors(fs).lightlevel Then
            markfloor = True
        Else
            ' same plane on both sides
            markfloor = False
        End If

        If worldhigh <> worldtop Or _
           sectors(bs).ceilingpic <> sectors(fs).ceilingpic Or _
           sectors(bs).lightlevel <> sectors(fs).lightlevel Then
            markceiling = True
        Else
            ' same plane on both sides
            markceiling = False
        End If

        If sectors(bs).ceilingheight <= sectors(fs).floorheight Or _
           sectors(bs).floorheight >= sectors(fs).ceilingheight Then
            ' closed door
            markceiling = True
            markfloor = True
        End If

        If worldhigh < worldtop Then
            ' top texture
            toptexture = texturetranslation(sides(sidedef).toptexture)
            If (lines(linedef).flags And ML_DONTPEGTOP) <> 0 Then
                ' top of texture at top
                rw_toptexturemid = worldtop
            Else
                vtop = sectors(bs).ceilingheight + textureheight(sides(sidedef).toptexture)
                ' bottom of texture
                rw_toptexturemid = vtop - viewz
            End If
        End If
        If worldlow > worldbottom Then
            ' bottom texture
            bottomtexture = texturetranslation(sides(sidedef).bottomtexture)
            If (lines(linedef).flags And ML_DONTPEGBOTTOM) <> 0 Then
                ' bottom of texture at bottom, top of texture at top
                rw_bottomtexturemid = worldtop
            Else
                ' top of texture at top
                rw_bottomtexturemid = worldlow
            End If
        End If
        rw_toptexturemid = rw_toptexturemid + sides(sidedef).rowoffset
        rw_bottomtexturemid = rw_bottomtexturemid + sides(sidedef).rowoffset

        ' allocate space for masked texture tables
        If sides(sidedef).midtexture <> 0 Then
            ' masked midtexture
            maskedtexture = True
            R_CheckOpenings rw_stopx - rw_x
            maskedtexturecol = lastopening - rw_x
            drawsegs(ds).maskedtexturecol = maskedtexturecol
            lastopening = lastopening + rw_stopx - rw_x
        End If
    End If

    ' calculate rw_offset (only needed for textured lines)
    segtextured = (midtexture <> 0) Or (toptexture <> 0) Or (bottomtexture <> 0) Or maskedtexture

    If segtextured Then
        offsetangle = rw_normalangle - rw_angle1

        If UGT(offsetangle, ANG180) Then offsetangle = -offsetangle

        If UGT(offsetangle, ANG90) Then offsetangle = ANG90

        sineval = finesine(FINEANGLE(offsetangle))
        rw_offset = FixedMul(hyp, sineval)

        If ULT(rw_normalangle - rw_angle1, ANG180) Then rw_offset = -rw_offset

        rw_offset = rw_offset + sides(sidedef).textureoffset + segs(curline).offset
        rw_centerangle = ANG90 + viewangle - rw_normalangle

        ' calculate light table
        '  use different light tables
        '  for horizontal / vertical / diagonal
        ' OPTIMIZE: get rid of LIGHTSEGSHIFT globally
        If fixedcolormap = 0 Then
            lightnum = (sectors(fs).lightlevel \ 16) + extralight

            If vertexes(segs(curline).v1).y = vertexes(segs(curline).v2).y Then
                lightnum = lightnum - 1
            ElseIf vertexes(segs(curline).v1).x = vertexes(segs(curline).v2).x Then
                lightnum = lightnum + 1
            End If

            If lightnum < 0 Then
                walllights = 0
            ElseIf lightnum >= LIGHTLEVELS Then
                walllights = LIGHTLEVELS - 1
            Else
                walllights = lightnum
            End If
        End If
    End If

    ' if a floor / ceiling plane is on the wrong side of the view plane, it is
    ' definitely invisible and doesn't need to be marked.
    If sectors(fs).floorheight >= viewz Then
        ' above view plane
        markfloor = False
    End If

    If sectors(fs).ceilingheight <= viewz And sectors(fs).ceilingpic <> skyflatnum Then
        ' below view plane
        markceiling = False
    End If

    ' calculate incremental stepping values for texture edges
    worldtop = (worldtop And &HFFFFFFF0) \ 16
    worldbottom = (worldbottom And &HFFFFFFF0) \ 16

    topstep = -FixedMul(rw_scalestep, worldtop)
    topfrac = (centeryfrac \ 16) - FixedMul(worldtop, rw_scale)

    bottomstep = -FixedMul(rw_scalestep, worldbottom)
    bottomfrac = (centeryfrac \ 16) - FixedMul(worldbottom, rw_scale)

    If bs <> -1 Then
        worldhigh = (worldhigh And &HFFFFFFF0) \ 16
        worldlow = (worldlow And &HFFFFFFF0) \ 16

        If worldhigh < worldtop Then
            pixhigh = (centeryfrac \ 16) - FixedMul(worldhigh, rw_scale)
            pixhighstep = -FixedMul(rw_scalestep, worldhigh)
        End If

        If worldlow > worldbottom Then
            pixlow = (centeryfrac \ 16) - FixedMul(worldlow, rw_scale)
            pixlowstep = -FixedMul(rw_scalestep, worldlow)
        End If
    End If

    ' render it
    If markceiling Then ceilingplane = R_CheckPlane(ceilingplane, rw_x, rw_stopx - 1)

    If markfloor Then floorplane = R_CheckPlane(floorplane, rw_x, rw_stopx - 1)

    R_RenderSegLoop

    ' save sprite clipping info
    If ((drawsegs(ds).silhouette And SIL_TOP) <> 0 Or maskedtexture) And drawsegs(ds).sprtopclip = -1 Then
        R_CheckOpenings rw_stopx - start
        CopyMemory openings(lastopening), ceilingclip(start), 4 * (rw_stopx - start)
        drawsegs(ds).sprtopclip = lastopening - start
        lastopening = lastopening + rw_stopx - start
    End If

    If ((drawsegs(ds).silhouette And SIL_BOTTOM) <> 0 Or maskedtexture) And drawsegs(ds).sprbottomclip = -1 Then
        R_CheckOpenings rw_stopx - start
        CopyMemory openings(lastopening), floorclip(start), 4 * (rw_stopx - start)
        drawsegs(ds).sprbottomclip = lastopening - start
        lastopening = lastopening + rw_stopx - start
    End If

    If maskedtexture And (drawsegs(ds).silhouette And SIL_TOP) = 0 Then
        drawsegs(ds).silhouette = drawsegs(ds).silhouette Or SIL_TOP
        drawsegs(ds).tsilheight = MININT
    End If
    If maskedtexture And (drawsegs(ds).silhouette And SIL_BOTTOM) = 0 Then
        drawsegs(ds).silhouette = drawsegs(ds).silhouette Or SIL_BOTTOM
        drawsegs(ds).bsilheight = MAXINT
    End If
    ds_p = ds_p + 1
End Sub
