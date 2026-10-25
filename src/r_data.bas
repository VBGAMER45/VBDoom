Attribute VB_Name = "r_data"
' r_data.bas - preparation of data for rendering, generation of lookups,
' caching, retrieval by name (r_data.c).
'
' Graphics.  DOOM graphics for walls and sprites is stored in vertical runs
' of opaque pixels (posts).  A column is composed of zero or more posts, a
' patch or sprite is composed of zero or more columns.
'
' Texture definition.  Each texture is composed of one or more patches,
' with patches being lumps stored in the WAD.  The lumps are referenced by
' number, and patched into the rectangular texture space using origin and
' possibly other attributes.
'
' Columns that are covered by one patch are drawn straight from the patch
' lump.  Columns covered by several patches are composited ("texture
' composite") into a block that is appended to wadbuf() at startup, so every
' column the drawers see is an offset into wadbuf().
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' texture_t
Private tex_name() As String
Private tex_width() As Long
Private tex_height() As Long
Private tex_patchcount() As Long
Private tex_firstpatch() As Long                 ' index into the tp_* arrays

' texpatch_t: a single patch from a texture definition, basically a
' rectangular area within the texture rectangle.
Private tp_originx() As Long
Private tp_originy() As Long
Private tp_patch() As Long                       ' lump number

Private texturewidthmask() As Long
Private tex_colbase() As Long                    ' first entry of the texture in collump()/colofs()
Private collump() As Long                        ' -1 = composite column
Private colofs() As Long                         ' patch column: wadbuf offset of the pixel data
                                                 ' composite column: offset inside the composite block
Private texturecompositesize() As Long
Private texturecomposite() As Long               ' wadbuf offset of the composite block
Private texgenerated() As Boolean
Private texindex As Collection                   ' upper case name -> first texture number

' R_DrawColumnInCache - clip and draw a column from a patch into a cached post.
Private Sub R_DrawColumnInCache(ByVal patch As Long, ByVal cache As Long, ByVal originy As Long, ByVal cacheheight As Long)
    Dim count As Long, position As Long, source As Long

    Do While wadbuf(patch) <> &HFF
        source = patch + 3
        count = wadbuf(patch + 1)
        position = originy + wadbuf(patch)

        If position < 0 Then
            count = count + position
            position = 0
        End If

        If position + count > cacheheight Then count = cacheheight - position

        If count > 0 Then CopyMemory wadbuf(cache + position), wadbuf(source), count

        patch = patch + wadbuf(patch + 1) + 4
    Loop
End Sub

' R_GenerateComposite - using the texture definition, the composite texture
' is created from the patches, and each column is cached.
Private Sub R_GenerateComposite(ByVal texnum As Long)
    Dim block As Long, i As Long, p As Long, realpatch As Long
    Dim x As Long, x1 As Long, x2 As Long, base As Long, patchcol As Long

    block = texturecomposite(texnum)
    base = tex_colbase(texnum)

    ' Composite the columns together.
    For i = 0 To tex_patchcount(texnum) - 1
        p = tex_firstpatch(texnum) + i
        realpatch = lumpinfo(tp_patch(p)).position
        x1 = tp_originx(p)
        x2 = x1 + RdShort(realpatch)

        If x1 < 0 Then x = 0 Else x = x1
        If x2 > tex_width(texnum) Then x2 = tex_width(texnum)

        Do While x < x2
            ' Column does not have multiple patches?
            If collump(base + x) < 0 Then
                patchcol = realpatch + RdLong(realpatch + 8 + (x - x1) * 4)
                R_DrawColumnInCache patchcol, block + colofs(base + x), tp_originy(p), tex_height(texnum)
            End If
            x = x + 1
        Loop
    Next

    texgenerated(texnum) = True
End Sub

' R_GenerateLookup
Private Sub R_GenerateLookup(ByVal texnum As Long)
    Dim patchcount() As Long, i As Long, p As Long, realpatch As Long
    Dim x As Long, x1 As Long, x2 As Long, base As Long, w As Long

    texturecompositesize(texnum) = 0
    base = tex_colbase(texnum)
    w = tex_width(texnum)

    ' Now count the number of columns that are covered by more than one
    ' patch.  Fill in the lump / offset, so columns with only a single patch
    ' are all done.
    ReDim patchcount(0 To w)

    For i = 0 To tex_patchcount(texnum) - 1
        p = tex_firstpatch(texnum) + i
        realpatch = lumpinfo(tp_patch(p)).position
        x1 = tp_originx(p)
        x2 = x1 + RdShort(realpatch)

        If x1 < 0 Then x = 0 Else x = x1
        If x2 > w Then x2 = w

        Do While x < x2
            patchcount(x) = patchcount(x) + 1
            collump(base + x) = tp_patch(p)
            colofs(base + x) = realpatch + RdLong(realpatch + 8 + (x - x1) * 4) + 3
            x = x + 1
        Loop
    Next

    For x = 0 To w - 1
        If patchcount(x) = 0 Then
            I_Printf "R_GenerateLookup: column without a patch (" & tex_name(texnum) & ")"
        End If
        If patchcount(x) <> 1 Then
            ' Use the cached block.
            collump(base + x) = -1
            colofs(base + x) = texturecompositesize(texnum)
            texturecompositesize(texnum) = texturecompositesize(texnum) + tex_height(texnum)
        End If
    Next
End Sub

' R_GetColumn - returns the wadbuf() offset of the column's pixel data.
Public Function R_GetColumn(ByVal tex As Long, ByVal col As Long) As Long
    Dim i As Long
    i = tex_colbase(tex) + (col And texturewidthmask(tex))
    If collump(i) >= 0 Then
        R_GetColumn = colofs(i)
        Exit Function
    End If
    If Not texgenerated(tex) Then R_GenerateComposite tex
    R_GetColumn = texturecomposite(tex) + colofs(i)
End Function

' True if the column is composited from several patches (no post data).
Public Function R_IsCompositeColumn(ByVal tex As Long, ByVal col As Long) As Boolean
    R_IsCompositeColumn = (collump(tex_colbase(tex) + (col And texturewidthmask(tex))) < 0)
End Function

Public Function R_TextureName(ByVal tex As Long) As String
    R_TextureName = tex_name(tex)
End Function

' R_InitTextures - initializes the texture list with the textures from the world map.
Private Sub R_InitTextures()
    Dim names As Long, nummappatches As Long, patchlookup() As Long
    Dim maptex As Long, maptex1 As Long, maptex2 As Long
    Dim numtextures1 As Long, numtextures2 As Long, maxoff As Long, maxoff2 As Long
    Dim directory As Long, offset As Long, mtexture As Long
    Dim i As Long, j As Long, totalwidth As Long, totalpatches As Long
    Dim mpatch As Long, p As Long, total As Long, wadbase As Long, nm As String

    ' Load the patch names from pnames.lmp.
    names = W_CacheLumpName("PNAMES")
    nummappatches = RdLong(names)
    ReDim patchlookup(0 To nummappatches)
    For i = 0 To nummappatches - 1
        patchlookup(i) = W_CheckNumForName(RdName(names + 4 + i * 8))
    Next

    ' Load the map texture definitions from textures.lmp.
    ' The data is contained in one or two lumps,
    '  TEXTURE1 for shareware, plus TEXTURE2 for commercial.
    maptex1 = W_CacheLumpName("TEXTURE1")
    numtextures1 = RdLong(maptex1)
    maxoff = W_LumpLength(W_GetNumForName("TEXTURE1"))

    If W_CheckNumForName("TEXTURE2") <> -1 Then
        maptex2 = W_CacheLumpName("TEXTURE2")
        numtextures2 = RdLong(maptex2)
        maxoff2 = W_LumpLength(W_GetNumForName("TEXTURE2"))
    Else
        maptex2 = -1
        numtextures2 = 0
        maxoff2 = 0
    End If
    numtextures = numtextures1 + numtextures2

    ReDim tex_name(0 To numtextures)
    ReDim tex_width(0 To numtextures)
    ReDim tex_height(0 To numtextures)
    ReDim tex_patchcount(0 To numtextures)
    ReDim tex_firstpatch(0 To numtextures)
    ReDim texturewidthmask(0 To numtextures)
    ReDim textureheight(0 To numtextures)
    ReDim tex_colbase(0 To numtextures)
    ReDim texturecompositesize(0 To numtextures)
    ReDim texturecomposite(0 To numtextures)
    ReDim texgenerated(0 To numtextures)
    ReDim tp_originx(0 To 1023)
    ReDim tp_originy(0 To 1023)
    ReDim tp_patch(0 To 1023)
    Set texindex = New Collection

    maptex = maptex1
    directory = maptex + 4
    totalwidth = 0
    totalpatches = 0

    For i = 0 To numtextures - 1
        If i = numtextures1 Then
            ' Start looking in second texture file.
            maptex = maptex2
            maxoff = maxoff2
            directory = maptex + 4
        End If

        offset = RdLong(directory)
        If offset > maxoff Then I_Error "R_InitTextures: bad texture directory"

        mtexture = maptex + offset

        tex_name(i) = RdName(mtexture)
        tex_width(i) = RdShort(mtexture + 12)
        tex_height(i) = RdShort(mtexture + 14)
        tex_patchcount(i) = RdShort(mtexture + 20)
        tex_firstpatch(i) = totalpatches

        If totalpatches + tex_patchcount(i) > UBound(tp_patch) Then
            ReDim Preserve tp_originx(0 To (totalpatches + tex_patchcount(i)) * 2)
            ReDim Preserve tp_originy(0 To (totalpatches + tex_patchcount(i)) * 2)
            ReDim Preserve tp_patch(0 To (totalpatches + tex_patchcount(i)) * 2)
        End If

        For j = 0 To tex_patchcount(i) - 1
            mpatch = mtexture + 22 + j * 10
            p = totalpatches + j
            tp_originx(p) = RdShort(mpatch)
            tp_originy(p) = RdShort(mpatch + 2)
            tp_patch(p) = patchlookup(RdShort(mpatch + 4))
            If tp_patch(p) = -1 Then
                I_Error "R_InitTextures: Missing patch in texture " & tex_name(i)
            End If
        Next
        totalpatches = totalpatches + tex_patchcount(i)

        tex_colbase(i) = totalwidth

        j = 1
        Do While j * 2 <= tex_width(i)
            j = j * 2
        Loop
        texturewidthmask(i) = j - 1
        textureheight(i) = tex_height(i) * FRACUNIT

        totalwidth = totalwidth + tex_width(i)

        nm = UCase$(tex_name(i))
        On Error Resume Next
        texindex.Add i, nm                       ' first texture of a name wins
        On Error GoTo 0

        directory = directory + 4
    Next

    ReDim collump(0 To totalwidth)
    ReDim colofs(0 To totalwidth)

    ' Precalculate whatever possible.
    For i = 0 To numtextures - 1
        R_GenerateLookup i
    Next

    ' Reserve space for every composite at the end of wadbuf().
    total = 0
    For i = 0 To numtextures - 1
        total = total + texturecompositesize(i)
    Next
    wadbase = UBound(wadbuf) + 1
    ReDim Preserve wadbuf(0 To wadbase + total + 1024)
    For i = 0 To numtextures - 1
        texturecomposite(i) = wadbase
        wadbase = wadbase + texturecompositesize(i)
    Next

    ' Create translation table for global animation.
    ReDim texturetranslation(0 To numtextures)
    For i = 0 To numtextures - 1
        texturetranslation(i) = i
    Next
End Sub

' R_InitFlats
Private Sub R_InitFlats()
    Dim i As Long
    firstflat = W_GetNumForName("F_START") + 1
    lastflat = W_GetNumForName("F_END") - 1
    numflats = lastflat - firstflat + 1

    ' Create translation table for global animation.
    ReDim flattranslation(0 To numflats)
    For i = 0 To numflats - 1
        flattranslation(i) = i
    Next
End Sub

' R_InitSpriteLumps - finds the width and hoffset of all sprites in the wad,
' so the sprite does not need to be cached completely just for having the
' header info ready during rendering.
Private Sub R_InitSpriteLumps()
    Dim i As Long, patch As Long

    firstspritelump = W_GetNumForName("S_START") + 1
    lastspritelump = W_GetNumForName("S_END") - 1

    numspritelumps = lastspritelump - firstspritelump + 1
    ReDim spritewidth(0 To numspritelumps)
    ReDim spriteoffset(0 To numspritelumps)
    ReDim spritetopoffset(0 To numspritelumps)

    For i = 0 To numspritelumps - 1
        patch = lumpinfo(firstspritelump + i).position
        If lumpinfo(firstspritelump + i).size >= 8 Then
            spritewidth(i) = RdShort(patch) * FRACUNIT
            spriteoffset(i) = RdShort(patch + 4) * FRACUNIT
            spritetopoffset(i) = RdShort(patch + 6) * FRACUNIT
        End If
    Next
End Sub

' R_InitColormaps
Private Sub R_InitColormaps()
    ' Load in the light tables.
    W_ReadLump W_GetNumForName("COLORMAP"), colormaps
End Sub

' R_InitData - locates all the lumps that will be used by all views.  Must
' be called after W_Init.
Public Sub R_InitData()
    R_InitTextures
    I_Printf "  InitTextures (" & numtextures & ")"
    R_InitFlats
    I_Printf "  InitFlats (" & numflats & ")"
    R_InitSpriteLumps
    I_Printf "  InitSprites (" & numspritelumps & ")"
    R_InitColormaps
    I_Printf "  InitColormaps"
End Sub

' R_FlatNumForName - retrieval, get a flat number for a flat name.
Public Function R_FlatNumForName(ByVal name As String) As Long
    Dim i As Long
    i = W_CheckNumForName(name)
    If i = -1 Then I_Error "R_FlatNumForName: " & name & " not found"
    R_FlatNumForName = i - firstflat
End Function

' R_CheckTextureNumForName - check whether texture is available.  Filter out
' NoTexture indicator.
Public Function R_CheckTextureNumForName(ByVal name As String) As Long
    Dim v As Variant
    ' "NoTexture" marker.
    If Left$(name, 1) = "-" Then
        R_CheckTextureNumForName = 0
        Exit Function
    End If
    On Error GoTo notfound
    v = texindex(UCase$(name))
    R_CheckTextureNumForName = v
    Exit Function
notfound:
    R_CheckTextureNumForName = -1
End Function

' R_TextureNumForName - calls R_CheckTextureNumForName, aborts with error
' message.
Public Function R_TextureNumForName(ByVal name As String) As Long
    Dim i As Long
    i = R_CheckTextureNumForName(name)
    If i = -1 Then I_Error "R_TextureNumForName: " & name & " not found"
    R_TextureNumForName = i
End Function

' R_PrecacheLevel - preloads all relevant graphics for the level.  Here the
' only work left is to build the composites of the wall textures in use.
Public Sub R_PrecacheLevel()
    Dim i As Long, present() As Boolean, t As Long

    ReDim present(0 To numtextures)
    For i = 0 To numsides - 1
        present(sides(i).toptexture) = True
        present(sides(i).midtexture) = True
        present(sides(i).bottomtexture) = True
    Next

    ' Sky texture is always present.  Note that F_SKY1 is the name used to
    ' indicate a sky floor/ceiling as a flat, while the sky texture is stored
    ' like a wall texture, with an episode dependend name.
    present(skytexture) = True

    For i = 0 To numtextures - 1
        If present(i) Then
            If Not texgenerated(i) Then
                If texturecompositesize(i) > 0 Then R_GenerateComposite i
            End If
        End If
    Next
End Sub
