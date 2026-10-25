Attribute VB_Name = "r_things"
' r_things.bas - refresh of things, i.e. objects represented by sprites (r_things.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const MINZ As Long = FRACUNIT * 4
Private Const BASEYCENTER As Long = 100

' A vissprite_t is a thing that will be drawn during a refresh.
Public Type vissprite_t
    x1 As Long
    x2 As Long
    gx As Long                                   ' for line side calculation
    gy As Long
    gz As Long                                   ' global bottom / top for silhouette clipping
    gzt As Long
    startfrac As Long                            ' horizontal position of x1
    vscale As Long
    xiscale As Long                              ' negative if flipped
    texturemid As Long
    patch As Long
    colormap As Long                             ' offset in colormaps(); -1 = shadow draw (C: NULL)
    mobjflags As Long
End Type

' Sprites are patches with a special naming convention so they can be
' recognized by R_InitSprites.  spriteframe_t data lives in the sf_* arrays,
' frame f of sprite s is at spr_firstframe(s) + f; rotation r at f * 8 + r.
Public numsprdefs As Long
Public spr_numframes() As Long
Public spr_firstframe() As Long
Public sf_rotate() As Long                       ' -1 unset, 0 false, 1 true
Public sf_lump() As Long                         ' (frame * 8 + rotation) -> lump - firstspritelump
Public sf_flip() As Byte
Private numframesalloc As Long

Public pspritescale As Long
Public pspriteiscale As Long

Public spritelights As Long                      ' scalelight row

' constant arrays used for psprite clipping and initializing clipping:
' negonearray / screenheightarray are openings() offsets (r_plane).

' variables used to look up and range check thing_t sprites patches
Private sprtemp_rotate(0 To 28) As Long
Private sprtemp_lump(0 To 28, 0 To 7) As Long
Private sprtemp_flip(0 To 28, 0 To 7) As Byte
Private maxframe As Long
Private spritename As String

' GAME FUNCTIONS
Public vissprites() As vissprite_t
Public vissprite_p As Long                       ' number of vissprites in use
Private maxvissprites As Long
Private vsorted() As Long                        ' sorted draw order (C: vsprsortedhead list)

Public mfloorclip As Long                        ' openings() offset
Public mceilingclip As Long
Public spryscale As Long
Public sprtopscreen As Long

' R_InstallSpriteLump - local function for R_InitSprites.
Private Sub R_InstallSpriteLump(ByVal lump As Long, ByVal frame As Long, ByVal rotation As Long, ByVal flipped As Boolean)
    Dim r As Long

    If frame >= 29 Or frame < 0 Or rotation > 8 Or rotation < 0 Then
        I_Error "R_InstallSpriteLump: Bad frame characters in lump " & lump
    End If

    If frame > maxframe Then maxframe = frame

    If rotation = 0 Then
        ' the lump should be used for all rotations
        If sprtemp_rotate(frame) = 0 Then
            I_Error "R_InitSprites: Sprite " & spritename & " frame " & Chr$(65 + frame) & " has multip rot=0 lump"
        End If
        If sprtemp_rotate(frame) = 1 Then
            I_Error "R_InitSprites: Sprite " & spritename & " frame " & Chr$(65 + frame) & " has rotations and a rot=0 lump"
        End If
        sprtemp_rotate(frame) = 0
        For r = 0 To 7
            sprtemp_lump(frame, r) = lump - firstspritelump
            If flipped Then sprtemp_flip(frame, r) = 1 Else sprtemp_flip(frame, r) = 0
        Next
        Exit Sub
    End If

    ' the lump is only used for one rotation
    If sprtemp_rotate(frame) = 0 Then
        I_Error "R_InitSprites: Sprite " & spritename & " frame " & Chr$(65 + frame) & " has rotations and a rot=0 lump"
    End If

    sprtemp_rotate(frame) = 1

    ' make 0 based
    rotation = rotation - 1
    If sprtemp_lump(frame, rotation) <> -1 Then
        I_Error "R_InitSprites: Sprite " & spritename & " : " & Chr$(65 + frame) & " : " & Chr$(49 + rotation) & " has two lumps mapped to it"
    End If

    sprtemp_lump(frame, rotation) = lump - firstspritelump
    If flipped Then sprtemp_flip(frame, rotation) = 1 Else sprtemp_flip(frame, rotation) = 0
End Sub

' R_InitSpriteDefs - pass a null terminated list of sprite names (4 chars
' exactly) to be used.  Builds the sprite rotation matrixes to account for
' horizontally flipped sprites.  Will report an error if the lumps are
' inconsistant.  Only called at startup.
'
' Sprite lump names are 4 characters for the actor, a letter for the frame,
' and a number for the rotation.  A sprite that is flippable will have an
' additional letter/number appended.  The rotation character can be 0 to
' signify no rotations.
Private Sub R_InitSpriteDefs()
    Dim i As Long, l As Long, frame As Long, rotation As Long, r As Long
    Dim start As Long, finish As Long, patched As Long, nm As String

    numsprdefs = NUMSPRITES
    If numsprdefs = 0 Then Exit Sub

    ReDim spr_numframes(0 To numsprdefs - 1)
    ReDim spr_firstframe(0 To numsprdefs - 1)
    numframesalloc = 0
    ReDim sf_rotate(0 To 255)
    ReDim sf_lump(0 To 256 * 8 - 1)
    ReDim sf_flip(0 To 256 * 8 - 1)

    start = firstspritelump - 1
    finish = lastspritelump + 1

    ' scan all the lump names for each of the names, noting the highest frame
    ' letter.  Just compare 4 characters as ints
    For i = 0 To numsprdefs - 1
        spritename = sprnames(i)
        For frame = 0 To 28
            sprtemp_rotate(frame) = -1
            For r = 0 To 7
                sprtemp_lump(frame, r) = -1
                sprtemp_flip(frame, r) = &HFF
            Next
        Next

        maxframe = -1

        ' scan the lumps, filling in the frames for whatever is found
        For l = start + 1 To finish - 1
            nm = lumpinfo(l).name
            If Left$(nm, 4) = spritename And Len(nm) >= 6 Then
                frame = Asc(Mid$(nm, 5, 1)) - 65
                rotation = Asc(Mid$(nm, 6, 1)) - 48

                If modifiedgame Then
                    patched = W_GetNumForName(nm)
                Else
                    patched = l
                End If

                R_InstallSpriteLump patched, frame, rotation, False

                If Len(nm) >= 8 Then
                    frame = Asc(Mid$(nm, 7, 1)) - 65
                    rotation = Asc(Mid$(nm, 8, 1)) - 48
                    R_InstallSpriteLump l, frame, rotation, True
                End If
            End If
        Next

        ' check the frames that were found for completeness
        If maxframe = -1 Then
            spr_numframes(i) = 0
        Else
            maxframe = maxframe + 1

            For frame = 0 To maxframe - 1
                Select Case sprtemp_rotate(frame)
                    Case -1
                        ' no rotations were found for that frame at all
                        I_Error "R_InitSprites: No patches found for " & sprnames(i) & " frame " & Chr$(frame + 65)
                    Case 0
                        ' only the first rotation is needed
                    Case 1
                        ' must have all 8 frames
                        For rotation = 0 To 7
                            If sprtemp_lump(frame, rotation) = -1 Then
                                I_Error "R_InitSprites: Sprite " & sprnames(i) & " frame " & Chr$(frame + 65) & " is missing rotations"
                            End If
                        Next
                End Select
            Next

            ' allocate space for the frames present and copy sprtemp to it
            spr_numframes(i) = maxframe
            spr_firstframe(i) = numframesalloc
            If numframesalloc + maxframe > UBound(sf_rotate) + 1 Then
                ReDim Preserve sf_rotate(0 To (numframesalloc + maxframe) * 2)
                ReDim Preserve sf_lump(0 To (numframesalloc + maxframe) * 2 * 8 + 7)
                ReDim Preserve sf_flip(0 To (numframesalloc + maxframe) * 2 * 8 + 7)
            End If
            For frame = 0 To maxframe - 1
                sf_rotate(numframesalloc + frame) = sprtemp_rotate(frame)
                For r = 0 To 7
                    sf_lump((numframesalloc + frame) * 8 + r) = sprtemp_lump(frame, r)
                    sf_flip((numframesalloc + frame) * 8 + r) = sprtemp_flip(frame, r)
                Next
            Next
            numframesalloc = numframesalloc + maxframe
        End If
    Next
End Sub

' R_InitSprites - called at program start.
Public Sub R_InitSprites()
    R_InitSpriteDefs
    maxvissprites = 256
    ReDim vissprites(0 To maxvissprites - 1)
    ReDim vsorted(0 To maxvissprites - 1)
End Sub

' R_ClearSprites - called at frame start.
Public Sub R_ClearSprites()
    vissprite_p = 0
End Sub

' R_NewVisSprite (C: MAXVISSPRITES = 128, extra sprites were dropped)
Private Function R_NewVisSprite() As Long
    If vissprite_p >= maxvissprites Then
        If maxvissprites >= 65536 Then I_Error "R_NewVisSprite: runaway sprite count (sector thing list loop?)"
        maxvissprites = maxvissprites * 2
        ReDim Preserve vissprites(0 To maxvissprites - 1)
        ReDim Preserve vsorted(0 To maxvissprites - 1)
    End If
    R_NewVisSprite = vissprite_p
    vissprite_p = vissprite_p + 1
End Function

' R_DrawMaskedColumn - used for sprites and masked mid textures.  Masked
' means: partly transparent, i.e. stored in posts/runs of opaque pixels.
' column is the offset of the post data in wadbuf().
Public Sub R_DrawMaskedColumn(ByVal column As Long)
    Dim topscreen As Long, bottomscreen As Long, basetexturemid As Long
    Dim topdelta As Long, length As Long, clip As Long

    basetexturemid = dc_texturemid

    topdelta = wadbuf(column)
    Do While topdelta <> &HFF
        length = wadbuf(column + 1)
        ' calculate unclipped screen coordinates for post
        topscreen = sprtopscreen + spryscale * topdelta
        bottomscreen = topscreen + spryscale * length

        dc_yl = ((topscreen + FRACUNIT - 1) And &HFFFF0000) \ &H10000
        dc_yh = ((bottomscreen - 1) And &HFFFF0000) \ &H10000

        clip = openings(mfloorclip + dc_x)
        If dc_yh >= clip Then dc_yh = clip - 1
        clip = openings(mceilingclip + dc_x)
        If dc_yl <= clip Then dc_yl = clip + 1

        If dc_yl <= dc_yh Then
            dc_source = column + 3
            dc_texturemid = basetexturemid - topdelta * FRACUNIT

            ' Drawn by either R_DrawColumn or (SHADOW) R_DrawFuzzColumn.
            R_ColFunc
        End If
        column = column + length + 4
        topdelta = wadbuf(column)
    Loop

    dc_texturemid = basetexturemid
End Sub

' Draws a composite (multi patch) texture column as one solid post.  The C
' code reinterpreted the raw composite as posts (the "medusa" effect).
Public Sub R_DrawSolidMaskedColumn(ByVal source As Long, ByVal texheight As Long)
    Dim topscreen As Long, bottomscreen As Long, clip As Long

    topscreen = sprtopscreen
    bottomscreen = topscreen + FixedMul(spryscale, texheight)

    dc_yl = ((topscreen + FRACUNIT - 1) And &HFFFF0000) \ &H10000
    dc_yh = ((bottomscreen - 1) And &HFFFF0000) \ &H10000

    clip = openings(mfloorclip + dc_x)
    If dc_yh >= clip Then dc_yh = clip - 1
    clip = openings(mceilingclip + dc_x)
    If dc_yl <= clip Then dc_yl = clip + 1

    If dc_yl <= dc_yh Then
        dc_source = source
        R_ColFunc
    End If
End Sub

' R_DrawVisSprite - mfloorclip and mceilingclip should also be set.
Private Sub R_DrawVisSprite(vis As vissprite_t)
    Dim texturecolumn As Long, frac As Long, patch As Long, pwidth As Long, x As Long

    patch = lumpinfo(vis.patch + firstspritelump).position
    pwidth = RdShort(patch)

    dc_colormap = vis.colormap

    If dc_colormap = -1 Then
        ' NULL colormap = shadow draw
        colfunc = COLF_FUZZ
    ElseIf (vis.mobjflags And MF_TRANSLATION) <> 0 Then
        colfunc = COLF_TRANS
        dc_translation = ((vis.mobjflags And MF_TRANSLATION) \ &H40000) - 256
    End If

    dc_iscale = SAR(Abs(vis.xiscale), detailshift)
    dc_texturemid = vis.texturemid
    frac = vis.startfrac
    spryscale = vis.vscale
    sprtopscreen = centeryfrac - FixedMul(dc_texturemid, spryscale)

    For x = vis.x1 To vis.x2
        dc_x = x
        texturecolumn = (frac And &HFFFF0000) \ &H10000
        If texturecolumn < 0 Or texturecolumn >= pwidth Then
            I_Error "R_DrawSpriteRange: bad texturecolumn"
        End If
        R_DrawMaskedColumn patch + RdLong(patch + 8 + texturecolumn * 4)
        frac = frac + vis.xiscale
    Next

    colfunc = COLF_BASE
End Sub

' R_ProjectSprite - generates a vissprite for a thing if it might be visible.
Private Sub R_ProjectSprite(ByVal thing As Long)
    Dim tr_x As Long, tr_y As Long, gxt As Long, gyt As Long, tx As Long, tz As Long
    Dim xscale As Long, x1 As Long, x2 As Long, sprite As Long, fr As Long
    Dim lump As Long, rot As Long, flip As Boolean, index As Long
    Dim vis As Long, ang As Long, iscale As Long

    ' transform the origin point
    tr_x = mobjs(thing).x - viewx
    tr_y = mobjs(thing).y - viewy

    gxt = FixedMul(tr_x, viewcos)
    gyt = -FixedMul(tr_y, viewsin)

    tz = gxt - gyt

    ' thing is behind view plane?
    If tz < MINZ Then Exit Sub

    xscale = FixedDiv(projection, tz)

    gxt = -FixedMul(tr_x, viewsin)
    gyt = FixedMul(tr_y, viewcos)
    tx = -(gyt + gxt)

    ' too far off the side?
    If Abs(tx) > tz * 4 Then Exit Sub

    ' decide which patch to use for sprite relative to player
    sprite = mobjs(thing).sprite
    If sprite < 0 Or sprite >= numsprdefs Then
        I_Error "R_ProjectSprite: invalid sprite number " & sprite
    End If
    fr = mobjs(thing).frame And FF_FRAMEMASK
    If fr >= spr_numframes(sprite) Then
        I_Error "R_ProjectSprite: invalid sprite frame " & sprite & " : " & mobjs(thing).frame
    End If
    fr = spr_firstframe(sprite) + fr

    If sf_rotate(fr) = 1 Then
        ' choose a different rotation based on player view
        ang = R_PointToAngle(mobjs(thing).x, mobjs(thing).y)
        rot = SHR(ang - mobjs(thing).angle + &H90000000, 29)
        lump = sf_lump(fr * 8 + rot)
        flip = (sf_flip(fr * 8 + rot) <> 0)
    Else
        ' use single rotation for all views
        lump = sf_lump(fr * 8)
        flip = (sf_flip(fr * 8) <> 0)
    End If

    ' calculate edges of the shape
    tx = tx - spriteoffset(lump)
    x1 = ((centerxfrac + FixedMul(tx, xscale)) And &HFFFF0000) \ &H10000

    ' off the right side?
    If x1 > viewwidth Then Exit Sub

    tx = tx + spritewidth(lump)
    x2 = (((centerxfrac + FixedMul(tx, xscale)) And &HFFFF0000) \ &H10000) - 1

    ' off the left side
    If x2 < 0 Then Exit Sub

    ' store information in a vissprite
    vis = R_NewVisSprite()
    vissprites(vis).mobjflags = mobjs(thing).flags
    vissprites(vis).vscale = xscale * pow2(detailshift)
    vissprites(vis).gx = mobjs(thing).x
    vissprites(vis).gy = mobjs(thing).y
    vissprites(vis).gz = mobjs(thing).z
    vissprites(vis).gzt = mobjs(thing).z + spritetopoffset(lump)
    vissprites(vis).texturemid = vissprites(vis).gzt - viewz
    If x1 < 0 Then vissprites(vis).x1 = 0 Else vissprites(vis).x1 = x1
    If x2 >= viewwidth Then vissprites(vis).x2 = viewwidth - 1 Else vissprites(vis).x2 = x2
    iscale = FixedDiv(FRACUNIT, xscale)

    If flip Then
        vissprites(vis).startfrac = spritewidth(lump) - 1
        vissprites(vis).xiscale = -iscale
    Else
        vissprites(vis).startfrac = 0
        vissprites(vis).xiscale = iscale
    End If

    If vissprites(vis).x1 > x1 Then
        vissprites(vis).startfrac = vissprites(vis).startfrac + vissprites(vis).xiscale * (vissprites(vis).x1 - x1)
    End If
    vissprites(vis).patch = lump

    ' get light level
    If (mobjs(thing).flags And MF_SHADOW) <> 0 Then
        ' shadow draw
        vissprites(vis).colormap = -1
    ElseIf fixedcolormap <> 0 Then
        ' fixed map
        vissprites(vis).colormap = fixedcolormap
    ElseIf (mobjs(thing).frame And FF_FULLBRIGHT) <> 0 Then
        ' full bright
        vissprites(vis).colormap = 0
    Else
        ' diminished light
        index = SAR(xscale, LIGHTSCALESHIFT - detailshift)
        If index >= MAXLIGHTSCALE Then index = MAXLIGHTSCALE - 1
        If index < 0 Then index = 0
        vissprites(vis).colormap = scalelight(spritelights, index)
    End If
End Sub

' R_AddSprites - during BSP traversal, this adds sprites by sector.
Public Sub R_AddSprites(ByVal sec As Long)
    Dim thing As Long, lightnum As Long

    ' BSP is traversed by subsector.  A sector might have been split into
    ' several subsectors during BSP building.  Thus we check whether its
    ' already added.
    If sectors(sec).validcount = validcount Then Exit Sub

    ' Well, now it will be done.
    sectors(sec).validcount = validcount

    lightnum = (sectors(sec).lightlevel \ 16) + extralight

    If lightnum < 0 Then
        spritelights = 0
    ElseIf lightnum >= LIGHTLEVELS Then
        spritelights = LIGHTLEVELS - 1
    Else
        spritelights = lightnum
    End If

    ' Handle all things in sector.
    thing = sectors(sec).thinglist
    Do While thing <> 0
        R_ProjectSprite thing
        thing = mobjs(thing).snext
    Loop
End Sub

' R_DrawPSprite
Private Sub R_DrawPSprite(ByVal player As Long, ByVal pspnum As Long)
    Dim tx As Long, x1 As Long, x2 As Long, st As Long, sprite As Long, fr As Long
    Dim lump As Long, flip As Boolean, vis As vissprite_t

    st = players(player).psprites(pspnum).state

    ' decide which patch to use
    sprite = states(st).sprite
    If sprite < 0 Or sprite >= numsprdefs Then
        I_Error "R_ProjectSprite: invalid sprite number " & sprite
    End If
    fr = states(st).frame And FF_FRAMEMASK
    If fr >= spr_numframes(sprite) Then
        I_Error "R_ProjectSprite: invalid sprite frame " & sprite & " : " & states(st).frame
    End If
    fr = spr_firstframe(sprite) + fr

    lump = sf_lump(fr * 8)
    flip = (sf_flip(fr * 8) <> 0)

    ' calculate edges of the shape
    tx = players(player).psprites(pspnum).sx - 160 * FRACUNIT

    tx = tx - spriteoffset(lump)
    x1 = ((centerxfrac + FixedMul(tx, pspritescale)) And &HFFFF0000) \ &H10000

    ' off the right side
    If x1 > viewwidth Then Exit Sub

    tx = tx + spritewidth(lump)
    x2 = (((centerxfrac + FixedMul(tx, pspritescale)) And &HFFFF0000) \ &H10000) - 1

    ' off the left side
    If x2 < 0 Then Exit Sub

    ' store information in a vissprite
    vis.mobjflags = 0
    vis.texturemid = (BASEYCENTER * FRACUNIT) + FRACUNIT \ 2 - (players(player).psprites(pspnum).sy - spritetopoffset(lump))
    If x1 < 0 Then vis.x1 = 0 Else vis.x1 = x1
    If x2 >= viewwidth Then vis.x2 = viewwidth - 1 Else vis.x2 = x2
    vis.vscale = pspritescale * pow2(detailshift)

    If flip Then
        vis.xiscale = -pspriteiscale
        vis.startfrac = spritewidth(lump) - 1
    Else
        vis.xiscale = pspriteiscale
        vis.startfrac = 0
    End If

    If vis.x1 > x1 Then vis.startfrac = vis.startfrac + vis.xiscale * (vis.x1 - x1)

    vis.patch = lump

    If players(player).powers(pw_invisibility) > 4 * 32 Or _
       (players(player).powers(pw_invisibility) And 8) <> 0 Then
        ' shadow draw
        vis.colormap = -1
    ElseIf fixedcolormap <> 0 Then
        ' fixed color
        vis.colormap = fixedcolormap
    ElseIf (states(st).frame And FF_FULLBRIGHT) <> 0 Then
        ' full bright
        vis.colormap = 0
    Else
        ' local light
        vis.colormap = scalelight(spritelights, MAXLIGHTSCALE - 1)
    End If

    R_DrawVisSprite vis
End Sub

' R_DrawPlayerSprites
Private Sub R_DrawPlayerSprites()
    Dim i As Long, lightnum As Long, mo As Long

    ' get light level
    mo = players(viewplayer).mo
    lightnum = (sectors(subsectors(mobjs(mo).subsector).sector).lightlevel \ 16) + extralight

    If lightnum < 0 Then
        spritelights = 0
    ElseIf lightnum >= LIGHTLEVELS Then
        spritelights = LIGHTLEVELS - 1
    Else
        spritelights = lightnum
    End If

    ' clip to screen bounds
    mfloorclip = SCREENHEIGHTARRAY
    mceilingclip = NEGONEARRAY

    ' add all active psprites
    For i = 0 To NUMPSPRITES - 1
        If players(viewplayer).psprites(i).state <> S_NULL Then R_DrawPSprite viewplayer, i
    Next
End Sub

' R_SortVisSprites - the C version repeatedly removes the smallest scale from
' the unsorted list (first one wins on ties), i.e. a stable ascending sort.
Private Sub R_SortVisSprites()
    Dim i As Long, j As Long, v As Long, s As Long
    For i = 0 To vissprite_p - 1
        v = i
        s = vissprites(v).vscale
        j = i - 1
        Do While j >= 0
            If vissprites(vsorted(j)).vscale <= s Then Exit Do
            vsorted(j + 1) = vsorted(j)
            j = j - 1
        Loop
        vsorted(j + 1) = v
    Next
End Sub

' R_DrawSprite
Private Sub R_DrawSprite(ByVal spr As Long)
    Dim ds As Long, x As Long, r1 As Long, r2 As Long
    Dim scl As Long, lowscale As Long, silhouette As Long
    Dim sx1 As Long, sx2 As Long, sscale As Long

    sx1 = vissprites(spr).x1
    sx2 = vissprites(spr).x2
    sscale = vissprites(spr).vscale

    For x = sx1 To sx2
        openings(CLIPBOTARRAY + x) = -2
        openings(CLIPTOPARRAY + x) = -2
    Next

    ' Scan drawsegs from end to start for obscuring segs.
    ' The first drawseg that has a greater scale is the clip seg.
    For ds = ds_p - 1 To 0 Step -1
        ' determine if the drawseg obscures the sprite
        If drawsegs(ds).x1 > sx2 Or drawsegs(ds).x2 < sx1 Or _
           (drawsegs(ds).silhouette = 0 And drawsegs(ds).maskedtexturecol = -1) Then
            ' does not cover sprite
            GoTo nextds
        End If

        If drawsegs(ds).x1 < sx1 Then r1 = sx1 Else r1 = drawsegs(ds).x1
        If drawsegs(ds).x2 > sx2 Then r2 = sx2 Else r2 = drawsegs(ds).x2

        If drawsegs(ds).scale1 > drawsegs(ds).scale2 Then
            lowscale = drawsegs(ds).scale2
            scl = drawsegs(ds).scale1
        Else
            lowscale = drawsegs(ds).scale1
            scl = drawsegs(ds).scale2
        End If

        If scl < sscale Then
            ' masked mid texture?
            If drawsegs(ds).maskedtexturecol <> -1 Then R_RenderMaskedSegRange ds, r1, r2
            ' seg is behind sprite
            GoTo nextds
        ElseIf lowscale < sscale Then
            If R_PointOnSegSide(vissprites(spr).gx, vissprites(spr).gy, drawsegs(ds).curline) = 0 Then
                If drawsegs(ds).maskedtexturecol <> -1 Then R_RenderMaskedSegRange ds, r1, r2
                GoTo nextds
            End If
        End If

        ' clip this piece of the sprite
        silhouette = drawsegs(ds).silhouette

        If vissprites(spr).gz >= drawsegs(ds).bsilheight Then silhouette = silhouette And (Not SIL_BOTTOM)

        If vissprites(spr).gzt <= drawsegs(ds).tsilheight Then silhouette = silhouette And (Not SIL_TOP)

        If silhouette = 1 Then
            ' bottom sil
            For x = r1 To r2
                If openings(CLIPBOTARRAY + x) = -2 Then
                    openings(CLIPBOTARRAY + x) = openings(drawsegs(ds).sprbottomclip + x)
                End If
            Next
        ElseIf silhouette = 2 Then
            ' top sil
            For x = r1 To r2
                If openings(CLIPTOPARRAY + x) = -2 Then
                    openings(CLIPTOPARRAY + x) = openings(drawsegs(ds).sprtopclip + x)
                End If
            Next
        ElseIf silhouette = 3 Then
            ' both
            For x = r1 To r2
                If openings(CLIPBOTARRAY + x) = -2 Then
                    openings(CLIPBOTARRAY + x) = openings(drawsegs(ds).sprbottomclip + x)
                End If
                If openings(CLIPTOPARRAY + x) = -2 Then
                    openings(CLIPTOPARRAY + x) = openings(drawsegs(ds).sprtopclip + x)
                End If
            Next
        End If
nextds:
    Next

    ' all clipping has been performed, so draw the sprite

    ' check for unclipped columns
    For x = sx1 To sx2
        If openings(CLIPBOTARRAY + x) = -2 Then openings(CLIPBOTARRAY + x) = viewheight
        If openings(CLIPTOPARRAY + x) = -2 Then openings(CLIPTOPARRAY + x) = -1
    Next

    mfloorclip = CLIPBOTARRAY
    mceilingclip = CLIPTOPARRAY
    R_DrawVisSprite vissprites(spr)
End Sub

' R_DrawMasked
Public Sub R_DrawMasked()
    Dim i As Long, ds As Long

    R_SortVisSprites

    ' draw all vissprites back to front
    For i = 0 To vissprite_p - 1
        R_DrawSprite vsorted(i)
    Next

    ' render any remaining masked mid textures
    For ds = ds_p - 1 To 0 Step -1
        If drawsegs(ds).maskedtexturecol <> -1 Then
            R_RenderMaskedSegRange ds, drawsegs(ds).x1, drawsegs(ds).x2
        End If
    Next

    ' draw the psprites on top of everything
    '  but does not draw on side views
    If viewangleoffset = 0 Then R_DrawPlayerSprites
End Sub
