Attribute VB_Name = "v_video"
' v_video.bas - frame buffer drawing: patches, blocks, rectangles (v_video.c).
'
' A "patch" is identified by its lump number.  The patch header is read from
' wadbuf(): width, height, leftoffset, topoffset (shorts) then width column
' offsets (longs), each column being a list of posts terminated by &HFF.
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public Sub V_Init()
    ReDim screens(0 To 5 * SCREENSIZE - 1)
End Sub

Public Function PatchWidth(ByVal lump As Long) As Long
    PatchWidth = RdShort(lumpinfo(lump).position)
End Function

Public Function PatchHeight(ByVal lump As Long) As Long
    PatchHeight = RdShort(lumpinfo(lump).position + 2)
End Function

Public Function PatchLeftOffset(ByVal lump As Long) As Long
    PatchLeftOffset = RdShort(lumpinfo(lump).position + 4)
End Function

Public Function PatchTopOffset(ByVal lump As Long) As Long
    PatchTopOffset = RdShort(lumpinfo(lump).position + 6)
End Function

Public Sub V_MarkRect(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
    M_AddToBox dirtybox, x, y
    M_AddToBox dirtybox, x + width - 1, y + height - 1
End Sub

Public Sub V_CopyRect(ByVal srcx As Long, ByVal srcy As Long, ByVal srcscrn As Long, _
                      ByVal width As Long, ByVal height As Long, _
                      ByVal destx As Long, ByVal desty As Long, ByVal destscrn As Long)
    Dim src As Long, dest As Long
    If srcx < 0 Or srcx + width > SCREENWIDTH Or srcy < 0 Or srcy + height > SCREENHEIGHT _
       Or destx < 0 Or destx + width > SCREENWIDTH Or desty < 0 Or desty + height > SCREENHEIGHT _
       Or srcscrn > 4 Or destscrn > 4 Then
        I_Error "Bad V_CopyRect"
    End If
    If width <= 0 Then Exit Sub
    V_MarkRect destx, desty, width, height
    src = srcscrn * SCREENSIZE + SCREENWIDTH * srcy + srcx
    dest = destscrn * SCREENSIZE + SCREENWIDTH * desty + destx
    Do While height > 0
        CopyMemory screens(dest), screens(src), width
        src = src + SCREENWIDTH
        dest = dest + SCREENWIDTH
        height = height - 1
    Loop
End Sub

' Masks a column based masked pic to the screen.  Pixels outside the screen
' are clipped (the C version aborts with "Bad V_DrawPatch" instead).
Public Sub V_DrawPatch(ByVal x As Long, ByVal y As Long, ByVal scrn As Long, ByVal lump As Long)
    DrawPatchInternal x, y, scrn, lump, False
End Sub

Public Sub V_DrawPatchDirect(ByVal x As Long, ByVal y As Long, ByVal scrn As Long, ByVal lump As Long)
    DrawPatchInternal x, y, scrn, lump, False
End Sub

' Masks a column based masked pic to the screen, flipped horizontally.
Public Sub V_DrawPatchFlipped(ByVal x As Long, ByVal y As Long, ByVal scrn As Long, ByVal lump As Long)
    DrawPatchInternal x, y, scrn, lump, True
End Sub

Private Sub DrawPatchInternal(ByVal x As Long, ByVal y As Long, ByVal scrn As Long, ByVal lump As Long, ByVal flipped As Boolean)
    Dim p As Long, w As Long, col As Long, colofs As Long, sx As Long
    Dim topdelta As Long, length As Long, src As Long, dy As Long, i As Long
    Dim base As Long

    If lump < 0 Then Exit Sub
    p = lumpinfo(lump).position
    y = y - RdShort(p + 6)
    x = x - RdShort(p + 4)
    w = RdShort(p)
    If scrn = 0 Then V_MarkRect x, y, w, RdShort(p + 2)
    base = scrn * SCREENSIZE

    For col = 0 To w - 1
        If flipped Then
            colofs = p + RdLong(p + 8 + (w - 1 - col) * 4)
        Else
            colofs = p + RdLong(p + 8 + col * 4)
        End If
        sx = x + col
        If sx >= 0 And sx < SCREENWIDTH Then
            topdelta = wadbuf(colofs)
            Do While topdelta <> &HFF
                length = wadbuf(colofs + 1)
                src = colofs + 3
                dy = y + topdelta
                For i = 0 To length - 1
                    If dy >= 0 And dy < SCREENHEIGHT Then
                        screens(base + dy * SCREENWIDTH + sx) = wadbuf(src + i)
                    End If
                    dy = dy + 1
                Next
                colofs = colofs + length + 4
                topdelta = wadbuf(colofs)
            Loop
        End If
    Next
End Sub

' Draw a linear block of pixels into the view buffer (src() starting at srcofs).
Public Sub V_DrawBlock(ByVal x As Long, ByVal y As Long, ByVal scrn As Long, _
                       ByVal width As Long, ByVal height As Long, src() As Byte, ByVal srcofs As Long)
    Dim dest As Long
    If x < 0 Or x + width > SCREENWIDTH Or y < 0 Or y + height > SCREENHEIGHT Or scrn > 4 Then
        I_Error "Bad V_DrawBlock"
    End If
    V_MarkRect x, y, width, height
    dest = scrn * SCREENSIZE + y * SCREENWIDTH + x
    Do While height > 0
        CopyMemory screens(dest), src(srcofs), width
        srcofs = srcofs + width
        dest = dest + SCREENWIDTH
        height = height - 1
    Loop
End Sub

' Gets a linear block of pixels from the view buffer into dest() (resized).
Public Sub V_GetBlock(ByVal x As Long, ByVal y As Long, ByVal scrn As Long, _
                      ByVal width As Long, ByVal height As Long, dest() As Byte)
    Dim src As Long, d As Long
    If x < 0 Or x + width > SCREENWIDTH Or y < 0 Or y + height > SCREENHEIGHT Or scrn > 4 Then
        I_Error "Bad V_GetBlock"
    End If
    ReDim dest(0 To width * height - 1)
    src = scrn * SCREENSIZE + y * SCREENWIDTH + x
    Do While height > 0
        CopyMemory dest(d), screens(src), width
        src = src + SCREENWIDTH
        d = d + width
        height = height - 1
    Loop
End Sub

' Fills a rectangle of screen scrn with a 64x64 flat (used for backgrounds).
Public Sub V_FillFlat(ByVal scrn As Long, ByVal flatlump As Long, ByVal x0 As Long, ByVal y0 As Long, _
                      ByVal width As Long, ByVal height As Long)
    Dim x As Long, y As Long, src As Long, dest As Long
    src = lumpinfo(flatlump).position
    For y = y0 To y0 + height - 1
        dest = scrn * SCREENSIZE + y * SCREENWIDTH
        For x = x0 To x0 + width - 1
            screens(dest + x) = wadbuf(src + ((y And 63) * 64) + (x And 63))
        Next
    Next
End Sub
