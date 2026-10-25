Attribute VB_Name = "p_spec"
' p_spec.bas - implements special effects: texture animation, height or
' lighting changes according to adjacent sectors, respective utility
' functions, etc.  Line Tag handling.  Line and Sector triggers (p_spec.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' Animating textures and planes.  There is another anim_t used in wi_stuff,
' unrelated.
Private Type anim_t
    istexture As Boolean
    picnum As Long
    basepic As Long
    numpics As Long
    speed As Long
End Type

Private Const MAXANIMS As Long = 32
Private anims(0 To 31) As anim_t
Private lastanim As Long

' Animating line specials (C: MAXLINEANIMS = 64)
Private linespeciallist() As Long
Private numlinespecials As Long

Private Const MAX_ADJOINING_SECTORS As Long = 20

Private zeroButton As button_t

' P_InitPicAnims.  Floor/ceiling animation sequences, defined by first and
' last frame, i.e. the flat (64x64 tile) name to be used.  The full
' animation sequence is given using all the flats between the start and end
' entry, in the order found in the WAD file.
Public Sub P_InitPicAnims()
    Dim defs As Variant, i As Long, istexture As Boolean, endname As String, startname As String

    defs = Array( _
        False, "NUKAGE3", "NUKAGE1", 8, _
        False, "FWATER4", "FWATER1", 8, _
        False, "SWATER4", "SWATER1", 8, _
        False, "LAVA4", "LAVA1", 8, _
        False, "BLOOD3", "BLOOD1", 8, _
        False, "RROCK08", "RROCK05", 8, _
        False, "SLIME04", "SLIME01", 8, _
        False, "SLIME08", "SLIME05", 8, _
        False, "SLIME12", "SLIME09", 8, _
        True, "BLODGR4", "BLODGR1", 8, _
        True, "SLADRIP3", "SLADRIP1", 8, _
        True, "BLODRIP4", "BLODRIP1", 8, _
        True, "FIREWALL", "FIREWALA", 8, _
        True, "GSTFONT3", "GSTFONT1", 8, _
        True, "FIRELAVA", "FIRELAV3", 8, _
        True, "FIREMAG3", "FIREMAG1", 8, _
        True, "FIREBLU2", "FIREBLU1", 8, _
        True, "ROCKRED3", "ROCKRED1", 8, _
        True, "BFALL4", "BFALL1", 8, _
        True, "SFALL4", "SFALL1", 8, _
        True, "WFALL4", "WFALL1", 8, _
        True, "DBRAIN4", "DBRAIN1", 8)

    ' Init animation
    lastanim = 0
    For i = 0 To (UBound(defs) + 1) \ 4 - 1
        istexture = defs(i * 4)
        endname = defs(i * 4 + 1)
        startname = defs(i * 4 + 2)

        If istexture Then
            ' different episode ?
            If R_CheckTextureNumForName(startname) = -1 Then GoTo nextdef

            anims(lastanim).picnum = R_TextureNumForName(endname)
            anims(lastanim).basepic = R_TextureNumForName(startname)
        Else
            If W_CheckNumForName(startname) = -1 Then GoTo nextdef

            anims(lastanim).picnum = R_FlatNumForName(endname)
            anims(lastanim).basepic = R_FlatNumForName(startname)
        End If

        anims(lastanim).istexture = istexture
        anims(lastanim).numpics = anims(lastanim).picnum - anims(lastanim).basepic + 1

        If anims(lastanim).numpics < 2 Then
            I_Error "P_InitPicAnims: bad cycle from " & startname & " to " & endname
        End If

        anims(lastanim).speed = defs(i * 4 + 3)
        lastanim = lastanim + 1
nextdef:
    Next
End Sub

' UTILITIES

' getSide - will return a side index given the number of the current
' sector, the line number, and the side (0/1) that you want.
Public Function getSide(ByVal currentSector As Long, ByVal line As Long, ByVal side As Long) As Long
    getSide = lines(linebuffer(sectors(currentSector).lines + line)).sidenum(side)
End Function

' getSector - will return a sector index given the number of the current
' sector, the line number and the side (0/1) that you want.
Public Function getSector(ByVal currentSector As Long, ByVal line As Long, ByVal side As Long) As Long
    getSector = sides(lines(linebuffer(sectors(currentSector).lines + line)).sidenum(side)).sector
End Function

' twoSided - given the sector number and the line number, it will tell you
' whether the line is two-sided or not.
Public Function twoSided(ByVal sector As Long, ByVal line As Long) As Long
    twoSided = lines(linebuffer(sectors(sector).lines + line)).flags And ML_TWOSIDED
End Function

' getNextSector - return sector index of the sector next to current.
' -1 if not two-sided line
Public Function getNextSector(ByVal line As Long, ByVal sec As Long) As Long
    If (lines(line).flags And ML_TWOSIDED) = 0 Then
        getNextSector = -1
        Exit Function
    End If

    If lines(line).frontsector = sec Then
        getNextSector = lines(line).backsector
    Else
        getNextSector = lines(line).frontsector
    End If
End Function

' P_FindLowestFloorSurrounding()
' FIND LOWEST FLOOR HEIGHT IN SURROUNDING SECTORS
Public Function P_FindLowestFloorSurrounding(ByVal sec As Long) As Long
    Dim i As Long, other As Long, floor As Long

    floor = sectors(sec).floorheight
    For i = 0 To sectors(sec).linecount - 1
        other = getNextSector(linebuffer(sectors(sec).lines + i), sec)
        If other <> -1 Then
            If sectors(other).floorheight < floor Then floor = sectors(other).floorheight
        End If
    Next
    P_FindLowestFloorSurrounding = floor
End Function

' P_FindHighestFloorSurrounding()
' FIND HIGHEST FLOOR HEIGHT IN SURROUNDING SECTORS
Public Function P_FindHighestFloorSurrounding(ByVal sec As Long) As Long
    Dim i As Long, other As Long, floor As Long

    floor = -500 * FRACUNIT
    For i = 0 To sectors(sec).linecount - 1
        other = getNextSector(linebuffer(sectors(sec).lines + i), sec)
        If other <> -1 Then
            If sectors(other).floorheight > floor Then floor = sectors(other).floorheight
        End If
    Next
    P_FindHighestFloorSurrounding = floor
End Function

' P_FindNextHighestFloor
' FIND NEXT HIGHEST FLOOR IN SURROUNDING SECTORS
' Note: this should be doable w/o a fixed array.
Public Function P_FindNextHighestFloor(ByVal sec As Long, ByVal currentheight As Long) As Long
    Dim i As Long, h As Long, min As Long, other As Long, height As Long
    Dim heightlist(0 To MAX_ADJOINING_SECTORS - 1) As Long

    height = currentheight
    h = 0
    For i = 0 To sectors(sec).linecount - 1
        other = getNextSector(linebuffer(sectors(sec).lines + i), sec)
        If other <> -1 Then
            If sectors(other).floorheight > height Then
                heightlist(h) = sectors(other).floorheight
                h = h + 1
            End If

            ' Check for overflow. Exit.
            If h >= MAX_ADJOINING_SECTORS Then
                I_Printf "Sector with more than 20 adjoining sectors"
                Exit For
            End If
        End If
    Next

    ' Find lowest height in list
    If h = 0 Then
        P_FindNextHighestFloor = currentheight
        Exit Function
    End If

    min = heightlist(0)

    ' Range checking?
    For i = 1 To h - 1
        If heightlist(i) < min Then min = heightlist(i)
    Next

    P_FindNextHighestFloor = min
End Function

' FIND LOWEST CEILING IN THE SURROUNDING SECTORS
Public Function P_FindLowestCeilingSurrounding(ByVal sec As Long) As Long
    Dim i As Long, other As Long, height As Long

    height = MAXINT
    For i = 0 To sectors(sec).linecount - 1
        other = getNextSector(linebuffer(sectors(sec).lines + i), sec)
        If other <> -1 Then
            If sectors(other).ceilingheight < height Then height = sectors(other).ceilingheight
        End If
    Next
    P_FindLowestCeilingSurrounding = height
End Function

' FIND HIGHEST CEILING IN THE SURROUNDING SECTORS
Public Function P_FindHighestCeilingSurrounding(ByVal sec As Long) As Long
    Dim i As Long, other As Long, height As Long

    height = 0
    For i = 0 To sectors(sec).linecount - 1
        other = getNextSector(linebuffer(sectors(sec).lines + i), sec)
        If other <> -1 Then
            If sectors(other).ceilingheight > height Then height = sectors(other).ceilingheight
        End If
    Next
    P_FindHighestCeilingSurrounding = height
End Function

' RETURN NEXT SECTOR # THAT LINE TAG REFERS TO
Public Function P_FindSectorFromLineTag(ByVal line As Long, ByVal start As Long) As Long
    Dim i As Long
    For i = start + 1 To numsectors - 1
        If sectors(i).tag = lines(line).tag Then
            P_FindSectorFromLineTag = i
            Exit Function
        End If
    Next
    P_FindSectorFromLineTag = -1
End Function

' Find minimum light from an adjacent sector
Public Function P_FindMinSurroundingLight(ByVal sector As Long, ByVal max As Long) As Long
    Dim i As Long, min As Long, check As Long

    min = max
    For i = 0 To sectors(sector).linecount - 1
        check = getNextSector(linebuffer(sectors(sector).lines + i), sector)
        If check <> -1 Then
            If sectors(check).lightlevel < min Then min = sectors(check).lightlevel
        End If
    Next
    P_FindMinSurroundingLight = min
End Function

' EVENTS
' Events are operations triggered by using, crossing, or shooting special
' lines, or by timed thinkers.

' P_CrossSpecialLine - TRIGGER.  Called every time a thing origin is about
' to cross a line with a non 0 special.
Public Sub P_CrossSpecialLine(ByVal linenum As Long, ByVal side As Long, ByVal thing As Long)
    Dim line As Long, ok As Boolean

    line = linenum

    ' Triggers that other things can activate
    If mobjs(thing).player < 0 Then
        ' Things that should NOT trigger specials...
        Select Case mobjs(thing).mtype
            Case MT_ROCKET, MT_PLASMA, MT_BFG, MT_TROOPSHOT, MT_HEADSHOT, MT_BRUISERSHOT
                Exit Sub
        End Select

        ok = False
        Select Case lines(line).special
            Case 39, 97, 125, 126, 4, 10, 88
                ' TELEPORT TRIGGER, TELEPORT RETRIGGER, TELEPORT MONSTERONLY TRIGGER,
                ' TELEPORT MONSTERONLY RETRIGGER, RAISE DOOR, PLAT DOWN-WAIT-UP-STAY,
                ' PLAT DOWN-WAIT-UP-STAY RETRIGGER
                ok = True
        End Select
        If Not ok Then Exit Sub
    End If

    ' Note: could use some const's here.
    Select Case lines(line).special
        ' TRIGGERS.
        ' All from here to RETRIGGERS.
        Case 2
            ' Open Door
            EV_DoDoor line, vld_open
            lines(line).special = 0
        Case 3
            ' Close Door
            EV_DoDoor line, vld_close
            lines(line).special = 0
        Case 4
            ' Raise Door
            EV_DoDoor line, vld_normal
            lines(line).special = 0
        Case 5
            ' Raise Floor
            EV_DoFloor line, raiseFloor
            lines(line).special = 0
        Case 6
            ' Fast Ceiling Crush & Raise
            EV_DoCeiling line, fastCrushAndRaise
            lines(line).special = 0
        Case 8
            ' Build Stairs
            EV_BuildStairs line, build8
            lines(line).special = 0
        Case 10
            ' PlatDownWaitUp
            EV_DoPlat line, downWaitUpStay, 0
            lines(line).special = 0
        Case 12
            ' Light Turn On - brightest near
            EV_LightTurnOn line, 0
            lines(line).special = 0
        Case 13
            ' Light Turn On 255
            EV_LightTurnOn line, 255
            lines(line).special = 0
        Case 16
            ' Close Door 30
            EV_DoDoor line, vld_close30ThenOpen
            lines(line).special = 0
        Case 17
            ' Start Light Strobing
            EV_StartLightStrobing line
            lines(line).special = 0
        Case 19
            ' Lower Floor
            EV_DoFloor line, lowerFloor
            lines(line).special = 0
        Case 22
            ' Raise floor to nearest height and change texture
            EV_DoPlat line, raiseToNearestAndChange, 0
            lines(line).special = 0
        Case 25
            ' Ceiling Crush and Raise
            EV_DoCeiling line, crushAndRaise
            lines(line).special = 0
        Case 30
            ' Raise floor to shortest texture height on either side of lines.
            EV_DoFloor line, raiseToTexture
            lines(line).special = 0
        Case 35
            ' Lights Very Dark
            EV_LightTurnOn line, 35
            lines(line).special = 0
        Case 36
            ' Lower Floor (TURBO)
            EV_DoFloor line, turboLower
            lines(line).special = 0
        Case 37
            ' LowerAndChange
            EV_DoFloor line, lowerAndChange
            lines(line).special = 0
        Case 38
            ' Lower Floor To Lowest
            EV_DoFloor line, lowerFloorToLowest
            lines(line).special = 0
        Case 39
            ' TELEPORT!
            EV_Teleport line, side, thing
            lines(line).special = 0
        Case 40
            ' RaiseCeilingLowerFloor
            EV_DoCeiling line, raiseToHighest
            EV_DoFloor line, lowerFloorToLowest
            lines(line).special = 0
        Case 44
            ' Ceiling Crush
            EV_DoCeiling line, lowerAndCrush
            lines(line).special = 0
        Case 52
            ' EXIT!
            G_ExitLevel
        Case 53
            ' Perpetual Platform Raise
            EV_DoPlat line, perpetualRaise, 0
            lines(line).special = 0
        Case 54
            ' Platform Stop
            EV_StopPlat line
            lines(line).special = 0
        Case 56
            ' Raise Floor Crush
            EV_DoFloor line, raiseFloorCrush
            lines(line).special = 0
        Case 57
            ' Ceiling Crush Stop
            EV_CeilingCrushStop line
            lines(line).special = 0
        Case 58
            ' Raise Floor 24
            EV_DoFloor line, raiseFloor24
            lines(line).special = 0
        Case 59
            ' Raise Floor 24 And Change
            EV_DoFloor line, raiseFloor24AndChange
            lines(line).special = 0
        Case 104
            ' Turn lights off in sector(tag)
            EV_TurnTagLightsOff line
            lines(line).special = 0
        Case 108
            ' Blazing Door Raise (faster than TURBO!)
            EV_DoDoor line, vld_blazeRaise
            lines(line).special = 0
        Case 109
            ' Blazing Door Open (faster than TURBO!)
            EV_DoDoor line, vld_blazeOpen
            lines(line).special = 0
        Case 100
            ' Build Stairs Turbo 16
            EV_BuildStairs line, turbo16
            lines(line).special = 0
        Case 110
            ' Blazing Door Close (faster than TURBO!)
            EV_DoDoor line, vld_blazeClose
            lines(line).special = 0
        Case 119
            ' Raise floor to nearest surr. floor
            EV_DoFloor line, raiseFloorToNearest
            lines(line).special = 0
        Case 121
            ' Blazing PlatDownWaitUpStay
            EV_DoPlat line, blazeDWUS, 0
            lines(line).special = 0
        Case 124
            ' Secret EXIT
            G_SecretExitLevel
        Case 125
            ' TELEPORT MonsterONLY
            If mobjs(thing).player < 0 Then
                EV_Teleport line, side, thing
                lines(line).special = 0
            End If
        Case 130
            ' Raise Floor Turbo
            EV_DoFloor line, raiseFloorTurbo
            lines(line).special = 0
        Case 141
            ' Silent Ceiling Crush & Raise
            EV_DoCeiling line, silentCrushAndRaise
            lines(line).special = 0

        ' RETRIGGERS.  All from here till end.
        Case 72
            ' Ceiling Crush
            EV_DoCeiling line, lowerAndCrush
        Case 73
            ' Ceiling Crush and Raise
            EV_DoCeiling line, crushAndRaise
        Case 74
            ' Ceiling Crush Stop
            EV_CeilingCrushStop line
        Case 75
            ' Close Door
            EV_DoDoor line, vld_close
        Case 76
            ' Close Door 30
            EV_DoDoor line, vld_close30ThenOpen
        Case 77
            ' Fast Ceiling Crush & Raise
            EV_DoCeiling line, fastCrushAndRaise
        Case 79
            ' Lights Very Dark
            EV_LightTurnOn line, 35
        Case 80
            ' Light Turn On - brightest near
            EV_LightTurnOn line, 0
        Case 81
            ' Light Turn On 255
            EV_LightTurnOn line, 255
        Case 82
            ' Lower Floor To Lowest
            EV_DoFloor line, lowerFloorToLowest
        Case 83
            ' Lower Floor
            EV_DoFloor line, lowerFloor
        Case 84
            ' LowerAndChange
            EV_DoFloor line, lowerAndChange
        Case 86
            ' Open Door
            EV_DoDoor line, vld_open
        Case 87
            ' Perpetual Platform Raise
            EV_DoPlat line, perpetualRaise, 0
        Case 88
            ' PlatDownWaitUp
            EV_DoPlat line, downWaitUpStay, 0
        Case 89
            ' Platform Stop
            EV_StopPlat line
        Case 90
            ' Raise Door
            EV_DoDoor line, vld_normal
        Case 91
            ' Raise Floor
            EV_DoFloor line, raiseFloor
        Case 92
            ' Raise Floor 24
            EV_DoFloor line, raiseFloor24
        Case 93
            ' Raise Floor 24 And Change
            EV_DoFloor line, raiseFloor24AndChange
        Case 94
            ' Raise Floor Crush
            EV_DoFloor line, raiseFloorCrush
        Case 95
            ' Raise floor to nearest height and change texture.
            EV_DoPlat line, raiseToNearestAndChange, 0
        Case 96
            ' Raise floor to shortest texture height on either side of lines.
            EV_DoFloor line, raiseToTexture
        Case 97
            ' TELEPORT!
            EV_Teleport line, side, thing
        Case 98
            ' Lower Floor (TURBO)
            EV_DoFloor line, turboLower
        Case 105
            ' Blazing Door Raise (faster than TURBO!)
            EV_DoDoor line, vld_blazeRaise
        Case 106
            ' Blazing Door Open (faster than TURBO!)
            EV_DoDoor line, vld_blazeOpen
        Case 107
            ' Blazing Door Close (faster than TURBO!)
            EV_DoDoor line, vld_blazeClose
        Case 120
            ' Blazing PlatDownWaitUpStay.
            EV_DoPlat line, blazeDWUS, 0
        Case 126
            ' TELEPORT MonsterONLY.
            If mobjs(thing).player < 0 Then EV_Teleport line, side, thing
        Case 128
            ' Raise To Nearest Floor
            EV_DoFloor line, raiseFloorToNearest
        Case 129
            ' Raise Floor Turbo
            EV_DoFloor line, raiseFloorTurbo
    End Select
End Sub

' P_ShootSpecialLine - IMPACT SPECIALS.  Called when a thing shoots a
' special line.
Public Sub P_ShootSpecialLine(ByVal thing As Long, ByVal line As Long)
    ' Impacts that other things can activate.
    If mobjs(thing).player < 0 Then
        If lines(line).special <> 46 Then Exit Sub ' OPEN DOOR IMPACT
    End If

    Select Case lines(line).special
        Case 24
            ' RAISE FLOOR
            EV_DoFloor line, raiseFloor
            P_ChangeSwitchTexture line, 0
        Case 46
            ' OPEN DOOR
            EV_DoDoor line, vld_open
            P_ChangeSwitchTexture line, 1
        Case 47
            ' RAISE FLOOR NEAR AND CHANGE
            EV_DoPlat line, raiseToNearestAndChange, 0
            P_ChangeSwitchTexture line, 0
    End Select
End Sub

' P_PlayerInSpecialSector - called every tic frame that the player origin is
' in a special sector
Public Sub P_PlayerInSpecialSector(ByVal player As Long)
    Dim sector As Long, mo As Long, hurt As Boolean

    mo = players(player).mo
    sector = subsectors(mobjs(mo).subsector).sector

    ' Falling, not all the way down yet?
    If mobjs(mo).z <> sectors(sector).floorheight Then Exit Sub

    ' Has hitten ground.
    Select Case sectors(sector).special
        Case 5
            ' HELLSLIME DAMAGE
            If players(player).powers(pw_ironfeet) = 0 Then
                If (leveltime And &H1F) = 0 Then P_DamageMobj mo, 0, 0, 10
            End If

        Case 7
            ' NUKAGE DAMAGE
            If players(player).powers(pw_ironfeet) = 0 Then
                If (leveltime And &H1F) = 0 Then P_DamageMobj mo, 0, 0, 5
            End If

        Case 16, 4
            ' SUPER HELLSLIME DAMAGE / STROBE HURT
            If players(player).powers(pw_ironfeet) = 0 Then
                hurt = True
            Else
                hurt = (P_Random() < 5)
            End If
            If hurt Then
                If (leveltime And &H1F) = 0 Then P_DamageMobj mo, 0, 0, 20
            End If

        Case 9
            ' SECRET SECTOR
            players(player).secretcount = players(player).secretcount + 1
            sectors(sector).special = 0

        Case 11
            ' EXIT SUPER DAMAGE! (for E1M8 finale)
            players(player).cheats = players(player).cheats And (Not CF_GODMODE)

            If (leveltime And &H1F) = 0 Then P_DamageMobj mo, 0, 0, 20

            If players(player).health <= 10 Then G_ExitLevel

        Case Else
            I_Error "P_PlayerInSpecialSector: unknown special " & sectors(sector).special
    End Select
End Sub

' P_UpdateSpecials - animate planes, scroll walls, etc.
Public Sub P_UpdateSpecials()
    Dim a As Long, pic As Long, i As Long, line As Long

    ' LEVEL TIMER
    If levelTimer Then
        levelTimeCount = levelTimeCount - 1
        If levelTimeCount = 0 Then G_ExitLevel
    End If

    ' ANIMATE FLATS AND TEXTURES GLOBALLY
    For a = 0 To lastanim - 1
        For i = anims(a).basepic To anims(a).basepic + anims(a).numpics - 1
            pic = anims(a).basepic + ((leveltime \ anims(a).speed + i) Mod anims(a).numpics)
            If anims(a).istexture Then
                texturetranslation(i) = pic
            Else
                flattranslation(i) = pic
            End If
        Next
    Next

    ' ANIMATE LINE SPECIALS
    For i = 0 To numlinespecials - 1
        line = linespeciallist(i)
        Select Case lines(line).special
            Case 48
                ' EFFECT FIRSTCOL SCROLL +
                sides(lines(line).sidenum(0)).textureoffset = sides(lines(line).sidenum(0)).textureoffset + FRACUNIT
        End Select
    Next

    ' DO BUTTONS
    For i = 0 To MAXBUTTONS - 1
        If buttonlist(i).btimer <> 0 Then
            buttonlist(i).btimer = buttonlist(i).btimer - 1
            If buttonlist(i).btimer = 0 Then
                Select Case buttonlist(i).where
                    Case bw_top
                        sides(lines(buttonlist(i).linenum).sidenum(0)).toptexture = buttonlist(i).btexture
                    Case bw_middle
                        sides(lines(buttonlist(i).linenum).sidenum(0)).midtexture = buttonlist(i).btexture
                    Case bw_bottom
                        sides(lines(buttonlist(i).linenum).sidenum(0)).bottomtexture = buttonlist(i).btexture
                End Select
                S_StartSound buttonlist(i).soundorg, sfx_swtchn
                buttonlist(i) = zeroButton
            End If
        End If
    Next
End Sub

' Special Stuff that can not be categorized
Public Function EV_DoDonut(ByVal line As Long) As Long
    Dim s1 As Long, s2 As Long, s3 As Long, secnum As Long, rtn As Long
    Dim i As Long, floor As Long, ln As Long

    secnum = -1
    rtn = 0
    Do
        secnum = P_FindSectorFromLineTag(line, secnum)
        If secnum < 0 Then Exit Do

        s1 = secnum

        ' ALREADY MOVING?  IF SO, KEEP GOING...
        If sectors(s1).specialdata <> 0 Then GoTo nextsec

        rtn = 1
        s2 = getNextSector(linebuffer(sectors(s1).lines), s1)
        If s2 = -1 Then GoTo nextsec             ' (C would crash here)
        For i = 0 To sectors(s2).linecount - 1
            ln = linebuffer(sectors(s2).lines + i)
            ' (C: "(!s2->lines[i]->flags & ML_TWOSIDED)" is always false)
            If lines(ln).backsector = s1 Then GoTo nextline
            s3 = lines(ln).backsector
            If s3 = -1 Then GoTo nextline        ' (C would crash here)

            ' Spawn rising slime
            floor = P_NewThinker()
            P_AddThinker floor
            sectors(s2).specialdata = floor
            thinkers(floor).func = TF_FLOOR
            floormoves(floor).typ = donutRaise
            floormoves(floor).crush = False
            floormoves(floor).direction = 1
            floormoves(floor).sector = s2
            floormoves(floor).speed = FLOORSPEED \ 2
            floormoves(floor).texture = sectors(s3).floorpic
            floormoves(floor).newspecial = 0
            floormoves(floor).floordestheight = sectors(s3).floorheight

            ' Spawn lowering donut-hole
            floor = P_NewThinker()
            P_AddThinker floor
            sectors(s1).specialdata = floor
            thinkers(floor).func = TF_FLOOR
            floormoves(floor).typ = lowerFloor
            floormoves(floor).crush = False
            floormoves(floor).direction = -1
            floormoves(floor).sector = s1
            floormoves(floor).speed = FLOORSPEED \ 2
            floormoves(floor).floordestheight = sectors(s3).floorheight
            Exit For
nextline:
        Next
nextsec:
    Loop
    EV_DoDonut = rtn
End Function

' SPECIAL SPAWNING

' P_SpawnSpecials - after the map has been loaded, scan for specials that
' spawn thinkers
Public Sub P_SpawnSpecials()
    Dim sector As Long, i As Long, tm As Long

    ' See if -TIMER needs to be used.
    levelTimer = False

    i = M_CheckParm("-avg")
    If i <> 0 And deathmatch <> 0 Then
        levelTimer = True
        levelTimeCount = 20& * 60 * 35
    End If

    i = M_CheckParm("-timer")
    If i <> 0 And deathmatch <> 0 And i < myargc - 1 Then
        tm = Val(myargv(i + 1)) * 60 * 35
        levelTimer = True
        levelTimeCount = tm
    End If

    ' Init special SECTORs.
    For sector = 0 To numsectors - 1
        If sectors(sector).special <> 0 Then
            Select Case sectors(sector).special
                Case 1
                    ' FLICKERING LIGHTS
                    P_SpawnLightFlash sector
                Case 2
                    ' STROBE FAST
                    P_SpawnStrobeFlash sector, FASTDARK, 0
                Case 3
                    ' STROBE SLOW
                    P_SpawnStrobeFlash sector, SLOWDARK, 0
                Case 4
                    ' STROBE FAST/DEATH SLIME
                    P_SpawnStrobeFlash sector, FASTDARK, 0
                    sectors(sector).special = 4
                Case 8
                    ' GLOWING LIGHT
                    P_SpawnGlowingLight sector
                Case 9
                    ' SECRET SECTOR
                    totalsecret = totalsecret + 1
                Case 10
                    ' DOOR CLOSE IN 30 SECONDS
                    P_SpawnDoorCloseIn30 sector
                Case 12
                    ' SYNC STROBE SLOW
                    P_SpawnStrobeFlash sector, SLOWDARK, 1
                Case 13
                    ' SYNC STROBE FAST
                    P_SpawnStrobeFlash sector, FASTDARK, 1
                Case 14
                    ' DOOR RAISE IN 5 MINUTES
                    P_SpawnDoorRaiseIn5Mins sector, sector
                Case 17
                    P_SpawnFireFlicker sector
            End Select
        End If
    Next

    ' Init line EFFECTs
    numlinespecials = 0
    ReDim linespeciallist(0 To 63)
    For i = 0 To numlines - 1
        Select Case lines(i).special
            Case 48
                ' EFFECT FIRSTCOL SCROLL+
                If numlinespecials > UBound(linespeciallist) Then
                    ReDim Preserve linespeciallist(0 To numlinespecials * 2)
                End If
                linespeciallist(numlinespecials) = i
                numlinespecials = numlinespecials + 1
        End Select
    Next

    ' Init other misc stuff
    For i = 0 To MAXCEILINGS - 1
        activeceilings(i) = 0
    Next

    For i = 0 To MAXPLATS - 1
        activeplats(i) = 0
    Next

    For i = 0 To MAXBUTTONS - 1
        buttonlist(i) = zeroButton
    Next
End Sub
