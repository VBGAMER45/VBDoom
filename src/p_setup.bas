Attribute VB_Name = "p_setup"
' p_setup.bas - do all the WAD I/O, get map description, set up initial
' state and misc. LUTs (p_setup.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' P_LoadVertexes
Private Sub P_LoadVertexes(ByVal lump As Long)
    Dim i As Long, ml As Long
    numvertexes = W_LumpLength(lump) \ 4
    ReDim vertexes(0 To numvertexes)
    ml = W_CacheLumpNum(lump)
    For i = 0 To numvertexes - 1
        vertexes(i).x = RdShort(ml) * FRACUNIT
        vertexes(i).y = RdShort(ml + 2) * FRACUNIT
        ml = ml + 4
    Next
End Sub

' P_LoadSegs
Private Sub P_LoadSegs(ByVal lump As Long)
    Dim i As Long, ml As Long, ldef As Long, side As Long

    numsegs = W_LumpLength(lump) \ 12
    ReDim segs(0 To numsegs)
    ml = W_CacheLumpNum(lump)
    For i = 0 To numsegs - 1
        segs(i).v1 = RdUShort(ml)
        segs(i).v2 = RdUShort(ml + 2)
        segs(i).angle = RdShort(ml + 4) * FRACUNIT
        segs(i).offset = RdShort(ml + 10) * FRACUNIT
        ldef = RdUShort(ml + 6)
        segs(i).linedef = ldef
        side = RdShort(ml + 8)
        segs(i).sidedef = lines(ldef).sidenum(side)
        segs(i).frontsector = sides(lines(ldef).sidenum(side)).sector
        If (lines(ldef).flags And ML_TWOSIDED) <> 0 Then
            segs(i).backsector = sides(lines(ldef).sidenum(side Xor 1)).sector
        Else
            segs(i).backsector = -1
        End If
        ml = ml + 12
    Next
End Sub

' P_LoadSubsectors
Private Sub P_LoadSubsectors(ByVal lump As Long)
    Dim i As Long, ms As Long
    numsubsectors = W_LumpLength(lump) \ 4
    ReDim subsectors(0 To numsubsectors)
    ms = W_CacheLumpNum(lump)
    For i = 0 To numsubsectors - 1
        subsectors(i).numlines = RdUShort(ms)
        subsectors(i).firstline = RdUShort(ms + 2)
        ms = ms + 4
    Next
End Sub

' P_LoadSectors
Private Sub P_LoadSectors(ByVal lump As Long)
    Dim i As Long, ms As Long
    numsectors = W_LumpLength(lump) \ 26
    ReDim sectors(0 To numsectors)
    ms = W_CacheLumpNum(lump)
    For i = 0 To numsectors - 1
        sectors(i).floorheight = RdShort(ms) * FRACUNIT
        sectors(i).ceilingheight = RdShort(ms + 2) * FRACUNIT
        sectors(i).floorpic = R_FlatNumForName(RdName(ms + 4))
        sectors(i).ceilingpic = R_FlatNumForName(RdName(ms + 12))
        sectors(i).lightlevel = RdShort(ms + 20)
        sectors(i).special = RdShort(ms + 22)
        sectors(i).tag = RdShort(ms + 24)
        sectors(i).thinglist = 0
        ms = ms + 26
    Next
End Sub

' P_LoadNodes
Private Sub P_LoadNodes(ByVal lump As Long)
    Dim i As Long, j As Long, k As Long, mn As Long
    numnodes = W_LumpLength(lump) \ 28
    ReDim nodes(0 To numnodes)
    mn = W_CacheLumpNum(lump)
    For i = 0 To numnodes - 1
        nodes(i).x = RdShort(mn) * FRACUNIT
        nodes(i).y = RdShort(mn + 2) * FRACUNIT
        nodes(i).dx = RdShort(mn + 4) * FRACUNIT
        nodes(i).dy = RdShort(mn + 6) * FRACUNIT
        For j = 0 To 1
            nodes(i).children(j) = RdUShort(mn + 24 + j * 2)
            For k = 0 To 3
                nodes(i).bbox(j, k) = RdShort(mn + 8 + j * 8 + k * 2) * FRACUNIT
            Next
        Next
        mn = mn + 28
    Next
End Sub

' P_LoadThings
Private Sub P_LoadThings(ByVal lump As Long)
    Dim i As Long, mt As Long, numthings As Long, spawn As Boolean
    Dim mthing As mapthing_t

    mt = W_CacheLumpNum(lump)
    numthings = W_LumpLength(lump) \ 10

    For i = 0 To numthings - 1
        mthing.x = RdShort(mt)
        mthing.y = RdShort(mt + 2)
        mthing.angle = RdShort(mt + 4)
        mthing.typ = RdShort(mt + 6)
        mthing.options = RdShort(mt + 8)
        mt = mt + 10

        spawn = True

        ' Do not spawn cool, new monsters if !commercial
        If gamemode <> commercial Then
            Select Case mthing.typ
                Case 68, 64, 88, 89, 69, 67, 71, 65, 66, 84
                    ' Arachnotron, Archvile, Boss Brain, Boss Shooter,
                    ' Hell Knight, Mancubus, Pain Elemental, Former Human
                    ' Commando, Revenant, Wolf SS
                    spawn = False
            End Select
        End If
        If spawn = False Then Exit For           ' (sic: C breaks out of the loop)

        ' Do spawn all other stuff.
        P_SpawnMapThing mthing
    Next
End Sub

' P_LoadLineDefs - also counts secret lines for intermissions.
Private Sub P_LoadLineDefs(ByVal lump As Long)
    Dim i As Long, mld As Long, v1 As Long, v2 As Long

    numlines = W_LumpLength(lump) \ 14
    ' lines(numlines) is a scratch line used by A_BossDeath / A_KeenDie
    ReDim lines(0 To numlines)
    mld = W_CacheLumpNum(lump)
    For i = 0 To numlines - 1
        lines(i).flags = RdShort(mld + 4)
        lines(i).special = RdShort(mld + 6)
        lines(i).tag = RdShort(mld + 8)
        v1 = RdUShort(mld)
        v2 = RdUShort(mld + 2)
        lines(i).v1 = v1
        lines(i).v2 = v2
        lines(i).dx = vertexes(v2).x - vertexes(v1).x
        lines(i).dy = vertexes(v2).y - vertexes(v1).y

        If lines(i).dx = 0 Then
            lines(i).slopetype = ST_VERTICAL
        ElseIf lines(i).dy = 0 Then
            lines(i).slopetype = ST_HORIZONTAL
        Else
            If FixedDiv(lines(i).dy, lines(i).dx) > 0 Then
                lines(i).slopetype = ST_POSITIVE
            Else
                lines(i).slopetype = ST_NEGATIVE
            End If
        End If

        If vertexes(v1).x < vertexes(v2).x Then
            lines(i).bbox(BOXLEFT) = vertexes(v1).x
            lines(i).bbox(BOXRIGHT) = vertexes(v2).x
        Else
            lines(i).bbox(BOXLEFT) = vertexes(v2).x
            lines(i).bbox(BOXRIGHT) = vertexes(v1).x
        End If

        If vertexes(v1).y < vertexes(v2).y Then
            lines(i).bbox(BOXBOTTOM) = vertexes(v1).y
            lines(i).bbox(BOXTOP) = vertexes(v2).y
        Else
            lines(i).bbox(BOXBOTTOM) = vertexes(v2).y
            lines(i).bbox(BOXTOP) = vertexes(v1).y
        End If

        lines(i).sidenum(0) = RdShort(mld + 10)
        lines(i).sidenum(1) = RdShort(mld + 12)

        If lines(i).sidenum(0) <> -1 Then
            lines(i).frontsector = sides(lines(i).sidenum(0)).sector
        Else
            lines(i).frontsector = -1
        End If

        If lines(i).sidenum(1) <> -1 Then
            lines(i).backsector = sides(lines(i).sidenum(1)).sector
        Else
            lines(i).backsector = -1
        End If
        mld = mld + 14
    Next
    lines(numlines).sidenum(0) = -1
    lines(numlines).sidenum(1) = -1
    lines(numlines).frontsector = -1
    lines(numlines).backsector = -1
End Sub

' P_LoadSideDefs
Private Sub P_LoadSideDefs(ByVal lump As Long)
    Dim i As Long, msd As Long
    numsides = W_LumpLength(lump) \ 30
    ReDim sides(0 To numsides)
    msd = W_CacheLumpNum(lump)
    For i = 0 To numsides - 1
        sides(i).textureoffset = RdShort(msd) * FRACUNIT
        sides(i).rowoffset = RdShort(msd + 2) * FRACUNIT
        sides(i).toptexture = R_TextureNumForName(RdName(msd + 4))
        sides(i).bottomtexture = R_TextureNumForName(RdName(msd + 12))
        sides(i).midtexture = R_TextureNumForName(RdName(msd + 20))
        sides(i).sector = RdShort(msd + 28)
        msd = msd + 30
    Next
End Sub

' P_LoadBlockMap
Private Sub P_LoadBlockMap(ByVal lump As Long)
    Dim i As Long, count As Long, p As Long

    count = W_LumpLength(lump) \ 2
    ReDim blockmaplump(0 To count)
    p = W_CacheLumpNum(lump)
    For i = 0 To count - 1
        blockmaplump(i) = RdShort(p + i * 2)
    Next
    blockmap = 4

    bmaporgx = blockmaplump(0) * FRACUNIT
    bmaporgy = blockmaplump(1) * FRACUNIT
    bmapwidth = blockmaplump(2)
    bmapheight = blockmaplump(3)

    ' clear out mobj chains
    ReDim blocklinks(0 To bmapwidth * bmapheight)
End Sub

' P_GroupLines - builds sector line lists and subsector sector numbers.
' Finds block bounding boxes for sectors.
Private Sub P_GroupLines()
    Dim i As Long, j As Long, total As Long, li As Long, sec As Long
    Dim bbox(0 To 3) As Long, block As Long, lb As Long

    ' look up sector number for each subsector
    For i = 0 To numsubsectors - 1
        subsectors(i).sector = sides(segs(subsectors(i).firstline).sidedef).sector
    Next

    ' count number of lines in each sector
    total = 0
    For i = 0 To numlines - 1
        total = total + 1
        sectors(lines(i).frontsector).linecount = sectors(lines(i).frontsector).linecount + 1

        If lines(i).backsector <> -1 And lines(i).backsector <> lines(i).frontsector Then
            sectors(lines(i).backsector).linecount = sectors(lines(i).backsector).linecount + 1
            total = total + 1
        End If
    Next

    ' build line tables for each sector
    ReDim linebuffer(0 To total)
    lb = 0
    For sec = 0 To numsectors - 1
        M_ClearBox bbox
        sectors(sec).lines = lb
        For li = 0 To numlines - 1
            If lines(li).frontsector = sec Or lines(li).backsector = sec Then
                linebuffer(lb) = li
                lb = lb + 1
                M_AddToBox bbox, vertexes(lines(li).v1).x, vertexes(lines(li).v1).y
                M_AddToBox bbox, vertexes(lines(li).v2).x, vertexes(lines(li).v2).y
            End If
        Next
        If lb - sectors(sec).lines <> sectors(sec).linecount Then I_Error "P_GroupLines: miscounted"

        ' set the degenmobj_t to the middle of the bounding box
        sectors(sec).soundorgx = (bbox(BOXRIGHT) + bbox(BOXLEFT)) \ 2
        sectors(sec).soundorgy = (bbox(BOXTOP) + bbox(BOXBOTTOM)) \ 2

        ' adjust bounding box to map blocks
        block = SAR(bbox(BOXTOP) - bmaporgy + MAXRADIUS, MAPBLOCKSHIFT)
        If block >= bmapheight Then block = bmapheight - 1
        sectors(sec).blockbox(BOXTOP) = block

        block = SAR(bbox(BOXBOTTOM) - bmaporgy - MAXRADIUS, MAPBLOCKSHIFT)
        If block < 0 Then block = 0
        sectors(sec).blockbox(BOXBOTTOM) = block

        block = SAR(bbox(BOXRIGHT) - bmaporgx + MAXRADIUS, MAPBLOCKSHIFT)
        If block >= bmapwidth Then block = bmapwidth - 1
        sectors(sec).blockbox(BOXRIGHT) = block

        block = SAR(bbox(BOXLEFT) - bmaporgx - MAXRADIUS, MAPBLOCKSHIFT)
        If block < 0 Then block = 0
        sectors(sec).blockbox(BOXLEFT) = block
    Next
End Sub

' P_LoadReject - (C: rejectmatrix points at the lump; a short lump is padded)
Private Sub P_LoadReject(ByVal lump As Long)
    Dim need As Long, n As Long
    need = (numsectors * numsectors + 7) \ 8 + 1
    n = W_LumpLength(lump)
    If n > need Then need = n
    ReDim rejectmatrix(0 To need)
    If n > 0 Then CopyMemory rejectmatrix(0), wadbuf(W_CacheLumpNum(lump)), n
End Sub

' P_SetupLevel
Public Sub P_SetupLevel(ByVal episode As Long, ByVal map As Long, ByVal playermask As Long, ByVal skill As Long)
    Dim i As Long, lumpname As String, lumpnum As Long

    totalkills = 0
    totalitems = 0
    totalsecret = 0
    wminfo.maxfrags = 0
    wminfo.partime = 180
    For i = 0 To MAXPLAYERS - 1
        players(i).killcount = 0
        players(i).secretcount = 0
        players(i).itemcount = 0
    Next

    ' Initial height of PointOfView will be set by player think.
    players(consoleplayer).viewz = 1

    ' Make sure all sounds are stopped before Z_FreeTags.
    S_Start

    P_InitThinkers

    ' find map name
    If gamemode = commercial Then
        lumpname = "map" & Format$(map, "00")
    Else
        lumpname = "E" & Chr$(48 + episode) & "M" & Chr$(48 + map)
    End If

    lumpnum = W_GetNumForName(lumpname)

    leveltime = 0

    ' note: most of this ordering is important
    P_LoadBlockMap lumpnum + ML_BLOCKMAP
    P_LoadVertexes lumpnum + ML_VERTEXES
    P_LoadSectors lumpnum + ML_SECTORS
    P_LoadSideDefs lumpnum + ML_SIDEDEFS

    P_LoadLineDefs lumpnum + ML_LINEDEFS
    P_LoadSubsectors lumpnum + ML_SSECTORS
    P_LoadNodes lumpnum + ML_NODES
    P_LoadSegs lumpnum + ML_SEGS

    P_LoadReject lumpnum + ML_REJECT

    P_GroupLines

    bodyqueslot = 0
    deathmatch_p = 0
    P_LoadThings lumpnum + ML_THINGS

    ' if deathmatch, randomly spawn the active players
    If deathmatch <> 0 Then
        For i = 0 To MAXPLAYERS - 1
            If playeringame(i) Then
                players(i).mo = 0
                G_DeathMatchSpawnPlayer i
            End If
        Next
    End If

    ' clear special respawning que
    iquehead = 0
    iquetail = 0

    ' set up world state
    P_SpawnSpecials

    ' preload graphics
    R_PrecacheLevel
End Sub

' P_Init
Public Sub P_Init()
    P_InitSwitchList
    P_InitPicAnims
    R_InitSprites
End Sub
