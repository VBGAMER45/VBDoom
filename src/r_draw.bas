Attribute VB_Name = "r_draw"
' r_draw.bas - the actual span/column drawing functions.  Here find the main
' potential for optimization, e.g. inline assembly, different algorithms
' (r_draw.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const SBARHEIGHT As Long = 32

' colfunc selection (C: function pointers colfunc / basecolfunc / fuzzcolfunc)
Public Const COLF_BASE As Long = 0               ' R_DrawColumn / R_DrawColumnLow
Public Const COLF_FUZZ As Long = 1               ' R_DrawFuzzColumn
Public Const COLF_TRANS As Long = 2              ' R_DrawTranslatedColumn
Public colfunc As Long

' All drawing to the view buffer is accomplished in this file.  The other
' refresh files only know about ccordinates, not the architecture of the
' frame buffer.  Conveniently, the frame buffer is a linear one, and we
' need only the base address, and the total size == width*height*depth/8.
Public ylookup(0 To 199) As Long                 ' offsets into screens() (screen 0)
Public columnofs(0 To 319) As Long

' Color tables for different players, translate a limited part to another
' (color ramps used for suit colors).
Public translationtables(0 To 767) As Byte

' first pixel in a column (possibly virtual)
Public dc_colormap As Long                       ' offset into colormaps()
Public dc_x As Long
Public dc_yl As Long
Public dc_yh As Long
Public dc_iscale As Long
Public dc_texturemid As Long
Public dc_source As Long                         ' offset into wadbuf()
Public dc_translation As Long                    ' offset into translationtables()

' spans
Public ds_y As Long
Public ds_x1 As Long
Public ds_x2 As Long
Public ds_colormap As Long
Public ds_xfrac As Long
Public ds_yfrac As Long
Public ds_xstep As Long
Public ds_ystep As Long
Public ds_source As Long                         ' start of a 64*64 tile image (wadbuf offset)

' Spectre/Invisibility.
Private Const FUZZTABLE As Long = 50
Private fuzzoffset(0 To 49) As Long
Private fuzzpos As Long

Public Sub R_ColFunc()
    Select Case colfunc
        Case COLF_BASE
            If detailshift = 0 Then R_DrawColumn Else R_DrawColumnLow
        Case COLF_FUZZ
            R_DrawFuzzColumn
        Case Else
            R_DrawTranslatedColumn
    End Select
End Sub

' A column is a vertical slice/span from a wall texture that, given the DOOM
' style restrictions on the view orientation, will always have constant z
' depth.  Thus a special case loop for very fast rendering can be used.  It
' has also been used with Wolfenstein 3D.
Public Sub R_DrawColumn()
    Dim count As Long, dest As Long, frac As Long, fracstep As Long

    count = dc_yh - dc_yl

    ' Zero length, column does not exceed a pixel.
    If count < 0 Then Exit Sub

    If dc_x < 0 Or dc_x >= SCREENWIDTH Or dc_yl < 0 Or dc_yh >= SCREENHEIGHT Then
        I_Error "R_DrawColumn: " & dc_yl & " to " & dc_yh & " at " & dc_x
    End If

    ' Framebuffer destination address.
    dest = ylookup(dc_yl) + columnofs(dc_x)

    ' Determine scaling, which is the only mapping to be done.
    fracstep = dc_iscale
    frac = dc_texturemid + (dc_yl - centery) * fracstep

    ' Inner loop that does the actual texture mapping, e.g. a DDA-lile scaling.
    ' This is as fast as it gets.
    Do
        ' Re-map color indices from wall texture column using a lighting/special
        ' effects LUT.
        screens(dest) = colormaps(dc_colormap + wadbuf(dc_source + ((frac And &H7F0000) \ &H10000)))
        dest = dest + SCREENWIDTH
        frac = frac + fracstep
        count = count - 1
    Loop While count >= 0
End Sub

' low detail: every column is drawn twice.  (The C version shifted dc_x in
' place, which broke two-tier walls; a local copy is used here.)
Public Sub R_DrawColumnLow()
    Dim count As Long, dest As Long, frac As Long, fracstep As Long, x As Long

    count = dc_yh - dc_yl

    ' Zero length.
    If count < 0 Then Exit Sub

    If dc_x < 0 Or dc_x >= SCREENWIDTH Or dc_yl < 0 Or dc_yh >= SCREENHEIGHT Then
        I_Error "R_DrawColumn: " & dc_yl & " to " & dc_yh & " at " & dc_x
    End If

    ' Blocky mode, need to multiply by 2.
    x = dc_x * 2
    dest = ylookup(dc_yl) + columnofs(x)
    fracstep = dc_iscale
    frac = dc_texturemid + (dc_yl - centery) * fracstep

    Do
        ' Hack. Does not work corretly.
        screens(dest) = colormaps(dc_colormap + wadbuf(dc_source + ((frac And &H7F0000) \ &H10000)))
        screens(dest + 1) = screens(dest)
        dest = dest + SCREENWIDTH
        frac = frac + fracstep
        count = count - 1
    Loop While count >= 0
End Sub

' Framebuffer postprocessing.  Creates a fuzzy image by copying pixels from
' adjacent ones to left and right.  Used with an all black colormap, this
' could create the SHADOW effect, i.e. spectres and invisible players.
Public Sub R_DrawFuzzColumn()
    Dim count As Long, dest As Long

    ' Adjust borders. Low...
    If dc_yl = 0 Then dc_yl = 1

    ' .. and high.
    If dc_yh = viewheight - 1 Then dc_yh = viewheight - 2

    count = dc_yh - dc_yl

    ' Zero length.
    If count < 0 Then Exit Sub

    If dc_x < 0 Or dc_x >= SCREENWIDTH Or dc_yl < 0 Or dc_yh >= SCREENHEIGHT Then
        I_Error "R_DrawFuzzColumn: " & dc_yl & " to " & dc_yh & " at " & dc_x
    End If

    dest = ylookup(dc_yl) + columnofs(dc_x)

    ' Looks like an attempt at dithering, using the colormap #6 (of 0-31, a
    ' bit brighter than average).
    Do
        ' Lookup framebuffer, and retrieve a pixel that is either one column
        ' left or right of the current one.  Add index from colormap to index.
        screens(dest) = colormaps(6 * 256 + screens(dest + fuzzoffset(fuzzpos)))

        ' Clamp table lookup index.
        fuzzpos = fuzzpos + 1
        If fuzzpos = FUZZTABLE Then fuzzpos = 0

        dest = dest + SCREENWIDTH
        count = count - 1
    Loop While count >= 0
End Sub

' R_DrawTranslatedColumn - used to draw player sprites with the green
' colorramp mapped to others.  Could be used with different translation
' tables, e.g. the lighter colored version of the BaronOfHell, the
' HellKnight, uses identical sprites, kinda brightened up.
Public Sub R_DrawTranslatedColumn()
    Dim count As Long, dest As Long, frac As Long, fracstep As Long

    count = dc_yh - dc_yl
    If count < 0 Then Exit Sub

    If dc_x < 0 Or dc_x >= SCREENWIDTH Or dc_yl < 0 Or dc_yh >= SCREENHEIGHT Then
        I_Error "R_DrawColumn: " & dc_yl & " to " & dc_yh & " at " & dc_x
    End If

    dest = ylookup(dc_yl) + columnofs(dc_x)

    ' Looks familiar.
    fracstep = dc_iscale
    frac = dc_texturemid + (dc_yl - centery) * fracstep

    ' Here we do an additional index re-mapping.
    Do
        ' Translation tables are used to map certain colorramps to other ones,
        ' used with PLAY sprites.  Thus the "green" ramp of the player 0 sprite
        ' is mapped to gray, red, black/indigo.
        screens(dest) = colormaps(dc_colormap + translationtables(dc_translation + _
                        wadbuf(dc_source + ((frac And &HFFFF0000) \ &H10000))))
        dest = dest + SCREENWIDTH
        frac = frac + fracstep
        count = count - 1
    Loop While count >= 0
End Sub

' R_InitTranslationTables - creates the translation tables to map the green
' color ramp to gray, brown, red.  Assumes a given structure of the PLAYPAL.
' Could be read from a lump instead.
Public Sub R_InitTranslationTables()
    Dim i As Long, v As Variant

    ' translate just the 16 green colors
    For i = 0 To 255
        If i >= &H70 And i <= &H7F Then
            ' map green ramp to gray, brown, red
            translationtables(i) = &H60 + (i And &HF)
            translationtables(i + 256) = &H40 + (i And &HF)
            translationtables(i + 512) = &H20 + (i And &HF)
        Else
            ' Keep all other colors as is.
            translationtables(i) = i
            translationtables(i + 256) = i
            translationtables(i + 512) = i
        End If
    Next

    v = Array(1, -1, 1, -1, 1, 1, -1, _
              1, 1, -1, 1, 1, 1, -1, _
              1, 1, 1, -1, -1, -1, -1, _
              1, -1, -1, 1, 1, 1, 1, -1, _
              1, -1, 1, 1, -1, -1, 1, _
              1, -1, -1, -1, -1, 1, 1, _
              1, 1, -1, 1, 1, -1, 1)
    For i = 0 To FUZZTABLE - 1
        fuzzoffset(i) = v(i) * SCREENWIDTH
    Next
End Sub

' R_DrawSpan - with DOOM style restrictions on view orientation, the floors
' and ceilings consist of horizontal slices or spans with constant z depth.
' However, rotation around the world z axis is possible, thus this mapping,
' while simpler and faster than perspective correct texture mapping, has to
' traverse the texture at an angle in all but a few cases.  In consequence,
' flats are not stored by column (like walls), and the inner loop has to
' step in texture space u and v.
Public Sub R_DrawSpan()
    Dim xfrac As Long, yfrac As Long, dest As Long, count As Long, spot As Long
    Dim xstep As Long, ystep As Long, src As Long, cmap As Long

    If ds_x2 < ds_x1 Or ds_x1 < 0 Or ds_x2 >= SCREENWIDTH Or ds_y < 0 Or ds_y > SCREENHEIGHT Then
        I_Error "R_DrawSpan: " & ds_x1 & " to " & ds_x2 & " at " & ds_y
    End If

    xfrac = ds_xfrac
    yfrac = ds_yfrac
    xstep = ds_xstep
    ystep = ds_ystep
    src = ds_source
    cmap = ds_colormap

    dest = ylookup(ds_y) + columnofs(ds_x1)

    ' We do not check for zero spans here?
    count = ds_x2 - ds_x1

    Do
        ' Current texture index in u,v.
        spot = ((yfrac And &H3F0000) \ 1024) + ((xfrac And &H3F0000) \ &H10000)

        ' Lookup pixel from flat texture tile, re-index using light/colormap.
        screens(dest) = colormaps(cmap + wadbuf(src + spot))
        dest = dest + 1

        ' Next step in u,v.
        xfrac = xfrac + xstep
        yfrac = yfrac + ystep
        count = count - 1
    Loop While count >= 0
End Sub

' Again..
Public Sub R_DrawSpanLow()
    Dim xfrac As Long, yfrac As Long, dest As Long, count As Long, spot As Long

    If ds_x2 < ds_x1 Or ds_x1 < 0 Or ds_x2 >= SCREENWIDTH Or ds_y < 0 Or ds_y > SCREENHEIGHT Then
        I_Error "R_DrawSpan: " & ds_x1 & " to " & ds_x2 & " at " & ds_y
    End If

    xfrac = ds_xfrac
    yfrac = ds_yfrac

    ' Blocky mode, need to multiply by 2.
    dest = ylookup(ds_y) + columnofs(ds_x1 * 2)

    count = ds_x2 - ds_x1
    Do
        spot = ((yfrac And &H3F0000) \ 1024) + ((xfrac And &H3F0000) \ &H10000)
        ' Lowres/blocky mode does it twice, while scale is adjusted appropriately.
        screens(dest) = colormaps(ds_colormap + wadbuf(ds_source + spot))
        screens(dest + 1) = screens(dest)
        dest = dest + 2

        xfrac = xfrac + ds_xstep
        yfrac = yfrac + ds_ystep
        count = count - 1
    Loop While count >= 0
End Sub

' R_InitBuffer - creats lookup tables that avoid multiplies and other
' hazzles for getting the framebuffer address of a pixel to draw.
Public Sub R_InitBuffer(ByVal w As Long, ByVal h As Long)
    Dim i As Long

    ' Handle resize, e.g. smaller view windows with border and/or status bar.
    viewwindowx = (SCREENWIDTH - w) \ 2

    ' Column offset. For windows.
    For i = 0 To w - 1
        columnofs(i) = viewwindowx + i
    Next

    ' Samw with base row offset.
    If w = SCREENWIDTH Then
        viewwindowy = 0
    Else
        viewwindowy = (SCREENHEIGHT - SBARHEIGHT - h) \ 2
    End If

    ' Preclaculate all row offsets.
    For i = 0 To h - 1
        ylookup(i) = (i + viewwindowy) * SCREENWIDTH
    Next
End Sub

' R_FillBackScreen - fills the back screen with a pattern for variable
' screen sizes.  Also draws a beveled edge.
Public Sub R_FillBackScreen()
    Dim src As Long, dest As Long, x As Long, y As Long, nm As String

    If scaledviewwidth = 320 Then Exit Sub

    If gamemode = commercial Then nm = "GRNROCK" Else nm = "FLOOR7_2"

    src = W_CacheLumpName(nm)
    dest = SCREENSIZE                            ' screens[1]

    For y = 0 To SCREENHEIGHT - SBARHEIGHT - 1
        For x = 0 To SCREENWIDTH \ 64 - 1
            CopyMemory screens(dest), wadbuf(src + ((y And 63) * 64)), 64
            dest = dest + 64
        Next
    Next

    For x = 0 To scaledviewwidth - 1 Step 8
        V_DrawPatch viewwindowx + x, viewwindowy - 8, 1, W_GetNumForName("brdr_t")
    Next
    For x = 0 To scaledviewwidth - 1 Step 8
        V_DrawPatch viewwindowx + x, viewwindowy + viewheight, 1, W_GetNumForName("brdr_b")
    Next
    For y = 0 To viewheight - 1 Step 8
        V_DrawPatch viewwindowx - 8, viewwindowy + y, 1, W_GetNumForName("brdr_l")
    Next
    For y = 0 To viewheight - 1 Step 8
        V_DrawPatch viewwindowx + scaledviewwidth, viewwindowy + y, 1, W_GetNumForName("brdr_r")
    Next

    ' Draw beveled edge.
    V_DrawPatch viewwindowx - 8, viewwindowy - 8, 1, W_GetNumForName("brdr_tl")
    V_DrawPatch viewwindowx + scaledviewwidth, viewwindowy - 8, 1, W_GetNumForName("brdr_tr")
    V_DrawPatch viewwindowx - 8, viewwindowy + viewheight, 1, W_GetNumForName("brdr_bl")
    V_DrawPatch viewwindowx + scaledviewwidth, viewwindowy + viewheight, 1, W_GetNumForName("brdr_br")
End Sub

' Copy a screen buffer.
Public Sub R_VideoErase(ByVal ofs As Long, ByVal count As Long)
    ' LFB copy.  This might not be a good idea if memcpy is not optiomal,
    ' e.g. byte by byte on a 32bit CPU, as GNU GCC/Linux libc did at one point.
    If count > 0 Then CopyMemory screens(ofs), screens(SCREENSIZE + ofs), count
End Sub

' R_DrawViewBorder - draws the border around the view for different size windows?
Public Sub R_DrawViewBorder()
    Dim top As Long, side As Long, ofs As Long, i As Long

    If scaledviewwidth = SCREENWIDTH Then Exit Sub

    top = ((SCREENHEIGHT - SBARHEIGHT) - viewheight) \ 2
    side = (SCREENWIDTH - scaledviewwidth) \ 2

    ' copy top and one line of left side
    R_VideoErase 0, top * SCREENWIDTH + side

    ' copy one line of right side and bottom
    ofs = (viewheight + top) * SCREENWIDTH - side
    R_VideoErase ofs, top * SCREENWIDTH + side

    ' copy sides using wraparound
    ofs = top * SCREENWIDTH + SCREENWIDTH - side
    side = side * 2

    For i = 1 To viewheight - 1
        R_VideoErase ofs, side
        ofs = ofs + SCREENWIDTH
    Next

    ' ?
    V_MarkRect 0, 0, SCREENWIDTH, SCREENHEIGHT - SBARHEIGHT
End Sub
