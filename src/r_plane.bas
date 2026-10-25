Attribute VB_Name = "r_plane"
' r_plane.bas - here is a core component: drawing the floors and ceilings,
' while maintaining a per column clipping list only.  Moreover, the sky
' areas have to be determined (r_plane.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' Now what is a visplane, anyway?  The top/bottom arrays live in vp_top() /
' vp_bottom() at plane * VPSTRIDE + x + 1 (x = -1 .. 320, the C pad bytes).
Public Type visplane_t
    height As Long
    picnum As Long
    lightlevel As Long
    minx As Long
    maxx As Long
End Type

Public Const VPSTRIDE As Long = 322

Public visplanes() As visplane_t
Public vp_top() As Byte
Public vp_bottom() As Byte
Public lastvisplane As Long                      ' number of visplanes in use
Public floorplane As Long                        ' -1 = none
Public ceilingplane As Long
Private maxvisplanes As Long                     ' C: 128 (overflow was fatal)

' openings(): every short* clip array of the C code is an offset into this.
'   NEGONEARRAY        negonearray       (-1 for every column)
'   SCREENHEIGHTARRAY  screenheightarray (viewheight for every column)
'   CLIPBOTARRAY / CLIPTOPARRAY  the local arrays of R_DrawSprite
'   OPENINGSSTART      start of the per frame openings
Public Const NEGONEARRAY As Long = 0
Public Const SCREENHEIGHTARRAY As Long = 320
Public Const CLIPBOTARRAY As Long = 640
Public Const CLIPTOPARRAY As Long = 960
Public Const OPENINGSSTART As Long = 1280
Public openings() As Long
Public lastopening As Long
Private maxopenings As Long

' Clip values are the solid pixel bounding the range.
'  floorclip starts out SCREENHEIGHT
'  ceilingclip starts out -1
Public floorclip(0 To 319) As Long
Public ceilingclip(0 To 319) As Long

' spanstart holds the start of a plane span, initialized to 0 at start
Private spanstart(0 To 199) As Long

' texture mapping
Public planezlight As Long                       ' zlight row
Public planeheight As Long

Public yslope(0 To 199) As Long
Public distscale(0 To 319) As Long
Public basexscale As Long
Public baseyscale As Long

Private cachedheight(0 To 199) As Long
Private cacheddistance(0 To 199) As Long
Private cachedxstep(0 To 199) As Long
Private cachedystep(0 To 199) As Long

' R_InitPlanes - only at game startup.
Public Sub R_InitPlanes()
    Dim i As Long
    maxvisplanes = 256
    ReDim visplanes(0 To maxvisplanes - 1)
    ReDim vp_top(0 To maxvisplanes * VPSTRIDE - 1)
    ReDim vp_bottom(0 To maxvisplanes * VPSTRIDE - 1)
    maxopenings = OPENINGSSTART + SCREENWIDTH * 64
    ReDim openings(0 To maxopenings - 1)
    For i = 0 To SCREENWIDTH - 1
        openings(NEGONEARRAY + i) = -1
        openings(SCREENHEIGHTARRAY + i) = viewheight
    Next
End Sub

' makes sure n more openings fit
Public Sub R_CheckOpenings(ByVal n As Long)
    If lastopening + n + SCREENWIDTH >= maxopenings Then
        If maxopenings >= 4194304 Then I_Error "R_CheckOpenings: runaway openings"
        maxopenings = maxopenings * 2
        ReDim Preserve openings(0 To maxopenings - 1)
    End If
End Sub

' R_MapPlane
' Uses global vars: planeheight, ds_source, basexscale, baseyscale, viewx, viewy
' BASIC PRIMITIVE
Private Sub R_MapPlane(ByVal y As Long, ByVal x1 As Long, ByVal x2 As Long)
    Dim angle As Long, distance As Long, length As Long, index As Long

    If x2 < x1 Or x1 < 0 Or x2 >= viewwidth Or y < 0 Or y > viewheight Then
        I_Error "R_MapPlane: " & x1 & ", " & x2 & " at " & y
    End If

    If planeheight <> cachedheight(y) Then
        cachedheight(y) = planeheight
        distance = FixedMul(planeheight, yslope(y))
        cacheddistance(y) = distance
        ds_xstep = FixedMul(distance, basexscale)
        cachedxstep(y) = ds_xstep
        ds_ystep = FixedMul(distance, baseyscale)
        cachedystep(y) = ds_ystep
    Else
        distance = cacheddistance(y)
        ds_xstep = cachedxstep(y)
        ds_ystep = cachedystep(y)
    End If

    length = FixedMul(distance, distscale(x1))
    angle = FINEANGLE(viewangle + xtoviewangle(x1))
    ds_xfrac = viewx + FixedMul(finecosine(angle), length)
    ds_yfrac = -viewy - FixedMul(finesine(angle), length)

    If fixedcolormap <> 0 Then
        ds_colormap = fixedcolormap
    Else
        index = SHR(distance, LIGHTZSHIFT)
        If index >= MAXLIGHTZ Then index = MAXLIGHTZ - 1
        ds_colormap = zlight(planezlight, index)
    End If

    ds_y = y
    ds_x1 = x1
    ds_x2 = x2

    ' high or low detail
    If detailshift = 0 Then R_DrawSpan Else R_DrawSpanLow
End Sub

' R_ClearPlanes - at begining of frame.
Public Sub R_ClearPlanes()
    Dim i As Long, angle As Long

    ' opening / clipping determination
    For i = 0 To viewwidth - 1
        floorclip(i) = viewheight
        ceilingclip(i) = -1
    Next

    lastvisplane = 0
    lastopening = OPENINGSSTART

    ' texture calculation
    For i = 0 To 199
        cachedheight(i) = 0
    Next

    ' left to right mapping
    angle = FINEANGLE(viewangle - ANG90)

    ' scale will be unit scale at SCREENWIDTH/2 distance
    basexscale = FixedDiv(finecosine(angle), centerxfrac)
    baseyscale = -FixedDiv(finesine(angle), centerxfrac)
End Sub

Private Sub GrowVisplanes()
    If maxvisplanes >= 65536 Then I_Error "R_FindPlane: runaway visplane count"
    maxvisplanes = maxvisplanes * 2
    ReDim Preserve visplanes(0 To maxvisplanes - 1)
    ReDim Preserve vp_top(0 To maxvisplanes * VPSTRIDE - 1)
    ReDim Preserve vp_bottom(0 To maxvisplanes * VPSTRIDE - 1)
End Sub

' R_FindPlane
Public Function R_FindPlane(ByVal height As Long, ByVal picnum As Long, ByVal lightlevel As Long) As Long
    Dim check As Long

    If picnum = skyflatnum Then
        height = 0                               ' all skys map together
        lightlevel = 0
    End If

    For check = 0 To lastvisplane - 1
        If height = visplanes(check).height Then
            If picnum = visplanes(check).picnum Then
                If lightlevel = visplanes(check).lightlevel Then
                    R_FindPlane = check
                    Exit Function
                End If
            End If
        End If
    Next

    If lastvisplane >= maxvisplanes Then GrowVisplanes

    check = lastvisplane
    lastvisplane = lastvisplane + 1

    visplanes(check).height = height
    visplanes(check).picnum = picnum
    visplanes(check).lightlevel = lightlevel
    visplanes(check).minx = SCREENWIDTH
    visplanes(check).maxx = -1

    FillMemory vp_top(check * VPSTRIDE), VPSTRIDE, &HFF

    R_FindPlane = check
End Function

' R_CheckPlane
Public Function R_CheckPlane(ByVal pl As Long, ByVal start As Long, ByVal stopx As Long) As Long
    Dim intrl As Long, intrh As Long, unionl As Long, unionh As Long, x As Long, base As Long

    If start < visplanes(pl).minx Then
        intrl = visplanes(pl).minx
        unionl = start
    Else
        unionl = visplanes(pl).minx
        intrl = start
    End If

    If stopx > visplanes(pl).maxx Then
        intrh = visplanes(pl).maxx
        unionh = stopx
    Else
        unionh = visplanes(pl).maxx
        intrh = stopx
    End If

    base = pl * VPSTRIDE + 1
    For x = intrl To intrh
        If vp_top(base + x) <> &HFF Then Exit For
    Next

    If x > intrh Then
        visplanes(pl).minx = unionl
        visplanes(pl).maxx = unionh

        ' use the same one
        R_CheckPlane = pl
        Exit Function
    End If

    ' make a new visplane
    If lastvisplane >= maxvisplanes Then GrowVisplanes
    visplanes(lastvisplane).height = visplanes(pl).height
    visplanes(lastvisplane).picnum = visplanes(pl).picnum
    visplanes(lastvisplane).lightlevel = visplanes(pl).lightlevel

    pl = lastvisplane
    lastvisplane = lastvisplane + 1
    visplanes(pl).minx = start
    visplanes(pl).maxx = stopx

    FillMemory vp_top(pl * VPSTRIDE), VPSTRIDE, &HFF

    R_CheckPlane = pl
End Function

' R_MakeSpans
Private Sub R_MakeSpans(ByVal x As Long, ByVal t1 As Long, ByVal b1 As Long, ByVal t2 As Long, ByVal b2 As Long)
    Do While t1 < t2 And t1 <= b1
        R_MapPlane t1, spanstart(t1), x - 1
        t1 = t1 + 1
    Loop
    Do While b1 > b2 And b1 >= t1
        R_MapPlane b1, spanstart(b1), x - 1
        b1 = b1 - 1
    Loop

    Do While t2 < t1 And t2 <= b2
        spanstart(t2) = x
        t2 = t2 + 1
    Loop
    Do While b2 > b1 And b2 >= t2
        spanstart(b2) = x
        b2 = b2 - 1
    Loop
End Sub

' R_DrawPlanes - at the end of each frame.
Public Sub R_DrawPlanes()
    Dim pl As Long, light As Long, x As Long, stopx As Long, angle As Long, base As Long

    For pl = 0 To lastvisplane - 1
        If visplanes(pl).minx <= visplanes(pl).maxx Then
            base = pl * VPSTRIDE + 1

            If visplanes(pl).picnum = skyflatnum Then
                ' sky flat
                ' Sky is allways drawn full bright, i.e. colormaps[0] is used.
                ' Because of this hack, sky is not affected by INVUL inverse mapping.
                dc_iscale = SAR(pspriteiscale, detailshift)
                dc_colormap = 0
                dc_texturemid = skytexturemid
                For x = visplanes(pl).minx To visplanes(pl).maxx
                    dc_yl = vp_top(base + x)
                    dc_yh = vp_bottom(base + x)

                    If dc_yl <= dc_yh Then
                        angle = SHR(viewangle + xtoviewangle(x), ANGLETOSKYSHIFT)
                        dc_x = x
                        dc_source = R_GetColumn(skytexture, angle)
                        R_ColFunc
                    End If
                Next
            Else
                ' regular flat
                ds_source = lumpinfo(firstflat + flattranslation(visplanes(pl).picnum)).position

                planeheight = Abs(visplanes(pl).height - viewz)
                light = (visplanes(pl).lightlevel \ 16) + extralight

                If light >= LIGHTLEVELS Then light = LIGHTLEVELS - 1
                If light < 0 Then light = 0

                planezlight = light

                vp_top(base + visplanes(pl).maxx + 1) = &HFF
                vp_top(base + visplanes(pl).minx - 1) = &HFF

                stopx = visplanes(pl).maxx + 1

                For x = visplanes(pl).minx To stopx
                    R_MakeSpans x, vp_top(base + x - 1), vp_bottom(base + x - 1), _
                                vp_top(base + x), vp_bottom(base + x)
                Next
            End If
        End If
    Next
End Sub
