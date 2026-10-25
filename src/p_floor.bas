Attribute VB_Name = "p_floor"
' p_floor.bas - floor animation: raising stairs (p_floor.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' FLOORS

' Move a plane (floor or ceiling) and check for crushing
Public Function T_MovePlane(ByVal sector As Long, ByVal speed As Long, ByVal dest As Long, _
                            ByVal crush As Boolean, ByVal floorOrCeiling As Long, ByVal direction As Long) As Long
    Dim flag As Boolean, lastpos As Long

    Select Case floorOrCeiling
        Case 0
            ' FLOOR
            Select Case direction
                Case -1
                    ' DOWN
                    If sectors(sector).floorheight - speed < dest Then
                        lastpos = sectors(sector).floorheight
                        sectors(sector).floorheight = dest
                        flag = P_ChangeSector(sector, crush)
                        If flag Then
                            sectors(sector).floorheight = lastpos
                            P_ChangeSector sector, crush
                            ' return crushed;
                        End If
                        T_MovePlane = res_pastdest
                        Exit Function
                    Else
                        lastpos = sectors(sector).floorheight
                        sectors(sector).floorheight = sectors(sector).floorheight - speed
                        flag = P_ChangeSector(sector, crush)
                        If flag Then
                            sectors(sector).floorheight = lastpos
                            P_ChangeSector sector, crush
                            T_MovePlane = res_crushed
                            Exit Function
                        End If
                    End If

                Case 1
                    ' UP
                    If sectors(sector).floorheight + speed > dest Then
                        lastpos = sectors(sector).floorheight
                        sectors(sector).floorheight = dest
                        flag = P_ChangeSector(sector, crush)
                        If flag Then
                            sectors(sector).floorheight = lastpos
                            P_ChangeSector sector, crush
                            ' return crushed;
                        End If
                        T_MovePlane = res_pastdest
                        Exit Function
                    Else
                        ' COULD GET CRUSHED
                        lastpos = sectors(sector).floorheight
                        sectors(sector).floorheight = sectors(sector).floorheight + speed
                        flag = P_ChangeSector(sector, crush)
                        If flag Then
                            If crush Then
                                T_MovePlane = res_crushed
                                Exit Function
                            End If
                            sectors(sector).floorheight = lastpos
                            P_ChangeSector sector, crush
                            T_MovePlane = res_crushed
                            Exit Function
                        End If
                    End If
            End Select

        Case 1
            ' CEILING
            Select Case direction
                Case -1
                    ' DOWN
                    If sectors(sector).ceilingheight - speed < dest Then
                        lastpos = sectors(sector).ceilingheight
                        sectors(sector).ceilingheight = dest
                        flag = P_ChangeSector(sector, crush)

                        If flag Then
                            sectors(sector).ceilingheight = lastpos
                            P_ChangeSector sector, crush
                            ' return crushed;
                        End If
                        T_MovePlane = res_pastdest
                        Exit Function
                    Else
                        ' COULD GET CRUSHED
                        lastpos = sectors(sector).ceilingheight
                        sectors(sector).ceilingheight = sectors(sector).ceilingheight - speed
                        flag = P_ChangeSector(sector, crush)

                        If flag Then
                            If crush Then
                                T_MovePlane = res_crushed
                                Exit Function
                            End If
                            sectors(sector).ceilingheight = lastpos
                            P_ChangeSector sector, crush
                            T_MovePlane = res_crushed
                            Exit Function
                        End If
                    End If

                Case 1
                    ' UP
                    If sectors(sector).ceilingheight + speed > dest Then
                        lastpos = sectors(sector).ceilingheight
                        sectors(sector).ceilingheight = dest
                        flag = P_ChangeSector(sector, crush)
                        If flag Then
                            sectors(sector).ceilingheight = lastpos
                            P_ChangeSector sector, crush
                            ' return crushed;
                        End If
                        T_MovePlane = res_pastdest
                        Exit Function
                    Else
                        lastpos = sectors(sector).ceilingheight
                        sectors(sector).ceilingheight = sectors(sector).ceilingheight + speed
                        flag = P_ChangeSector(sector, crush)
                    End If
            End Select
    End Select
    T_MovePlane = res_ok
End Function

' MOVE A FLOOR TO IT'S DESTINATION (UP OR DOWN)
Public Sub T_MoveFloor(ByVal floor As Long)
    Dim res As Long, sec As Long

    sec = floormoves(floor).sector
    res = T_MovePlane(sec, floormoves(floor).speed, floormoves(floor).floordestheight, _
                      floormoves(floor).crush, 0, floormoves(floor).direction)

    If (leveltime And 7) = 0 Then S_StartSound SectorSoundOrg(sec), sfx_stnmov

    If res = res_pastdest Then
        sectors(sec).specialdata = 0

        If floormoves(floor).direction = 1 Then
            Select Case floormoves(floor).typ
                Case donutRaise
                    sectors(sec).special = floormoves(floor).newspecial
                    sectors(sec).floorpic = floormoves(floor).texture
            End Select
        ElseIf floormoves(floor).direction = -1 Then
            Select Case floormoves(floor).typ
                Case lowerAndChange
                    sectors(sec).special = floormoves(floor).newspecial
                    sectors(sec).floorpic = floormoves(floor).texture
            End Select
        End If
        P_RemoveThinker floor

        S_StartSound SectorSoundOrg(sec), sfx_pstop
    End If
End Sub

' HANDLE FLOOR TYPES
Public Function EV_DoFloor(ByVal line As Long, ByVal floortype As Long) As Long
    Dim secnum As Long, rtn As Long, i As Long, sec As Long, floor As Long
    Dim minsize As Long, side As Long, other As Long

    secnum = -1
    rtn = 0
    Do
        secnum = P_FindSectorFromLineTag(line, secnum)
        If secnum < 0 Then Exit Do

        sec = secnum

        ' ALREADY MOVING?  IF SO, KEEP GOING...
        If sectors(sec).specialdata <> 0 Then GoTo nextsec

        ' new floor thinker
        rtn = 1
        floor = P_NewThinker()
        P_AddThinker floor
        sectors(sec).specialdata = floor
        thinkers(floor).func = TF_FLOOR
        floormoves(floor).typ = floortype
        floormoves(floor).crush = False

        Select Case floortype
            Case lowerFloor
                floormoves(floor).direction = -1
                floormoves(floor).sector = sec
                floormoves(floor).speed = FLOORSPEED
                floormoves(floor).floordestheight = P_FindHighestFloorSurrounding(sec)

            Case lowerFloorToLowest
                floormoves(floor).direction = -1
                floormoves(floor).sector = sec
                floormoves(floor).speed = FLOORSPEED
                floormoves(floor).floordestheight = P_FindLowestFloorSurrounding(sec)

            Case turboLower
                floormoves(floor).direction = -1
                floormoves(floor).sector = sec
                floormoves(floor).speed = FLOORSPEED * 4
                floormoves(floor).floordestheight = P_FindHighestFloorSurrounding(sec)
                If floormoves(floor).floordestheight <> sectors(sec).floorheight Then
                    floormoves(floor).floordestheight = floormoves(floor).floordestheight + 8 * FRACUNIT
                End If

            Case raiseFloorCrush, raiseFloor
                If floortype = raiseFloorCrush Then floormoves(floor).crush = True
                floormoves(floor).direction = 1
                floormoves(floor).sector = sec
                floormoves(floor).speed = FLOORSPEED
                floormoves(floor).floordestheight = P_FindLowestCeilingSurrounding(sec)
                If floormoves(floor).floordestheight > sectors(sec).ceilingheight Then
                    floormoves(floor).floordestheight = sectors(sec).ceilingheight
                End If
                If floortype = raiseFloorCrush Then
                    floormoves(floor).floordestheight = floormoves(floor).floordestheight - 8 * FRACUNIT
                End If

            Case raiseFloorTurbo
                floormoves(floor).direction = 1
                floormoves(floor).sector = sec
                floormoves(floor).speed = FLOORSPEED * 4
                floormoves(floor).floordestheight = P_FindNextHighestFloor(sec, sectors(sec).floorheight)

            Case raiseFloorToNearest
                floormoves(floor).direction = 1
                floormoves(floor).sector = sec
                floormoves(floor).speed = FLOORSPEED
                floormoves(floor).floordestheight = P_FindNextHighestFloor(sec, sectors(sec).floorheight)

            Case raiseFloor24
                floormoves(floor).direction = 1
                floormoves(floor).sector = sec
                floormoves(floor).speed = FLOORSPEED
                floormoves(floor).floordestheight = sectors(sec).floorheight + 24 * FRACUNIT

            Case raiseFloor512
                floormoves(floor).direction = 1
                floormoves(floor).sector = sec
                floormoves(floor).speed = FLOORSPEED
                floormoves(floor).floordestheight = sectors(sec).floorheight + 512 * FRACUNIT

            Case raiseFloor24AndChange
                floormoves(floor).direction = 1
                floormoves(floor).sector = sec
                floormoves(floor).speed = FLOORSPEED
                floormoves(floor).floordestheight = sectors(sec).floorheight + 24 * FRACUNIT
                sectors(sec).floorpic = sectors(lines(line).frontsector).floorpic
                sectors(sec).special = sectors(lines(line).frontsector).special

            Case raiseToTexture
                minsize = MAXINT
                floormoves(floor).direction = 1
                floormoves(floor).sector = sec
                floormoves(floor).speed = FLOORSPEED
                For i = 0 To sectors(sec).linecount - 1
                    If twoSided(secnum, i) <> 0 Then
                        side = getSide(secnum, i, 0)
                        If sides(side).bottomtexture >= 0 Then
                            If textureheight(sides(side).bottomtexture) < minsize Then
                                minsize = textureheight(sides(side).bottomtexture)
                            End If
                        End If
                        side = getSide(secnum, i, 1)
                        If sides(side).bottomtexture >= 0 Then
                            If textureheight(sides(side).bottomtexture) < minsize Then
                                minsize = textureheight(sides(side).bottomtexture)
                            End If
                        End If
                    End If
                Next
                floormoves(floor).floordestheight = sectors(sec).floorheight + minsize

            Case lowerAndChange
                floormoves(floor).direction = -1
                floormoves(floor).sector = sec
                floormoves(floor).speed = FLOORSPEED
                floormoves(floor).floordestheight = P_FindLowestFloorSurrounding(sec)
                floormoves(floor).texture = sectors(sec).floorpic

                For i = 0 To sectors(sec).linecount - 1
                    If twoSided(secnum, i) <> 0 Then
                        If sides(getSide(secnum, i, 0)).sector = secnum Then
                            other = getSector(secnum, i, 1)
                        Else
                            other = getSector(secnum, i, 0)
                        End If
                        If sectors(other).floorheight = floormoves(floor).floordestheight Then
                            floormoves(floor).texture = sectors(other).floorpic
                            floormoves(floor).newspecial = sectors(other).special
                            Exit For
                        End If
                    End If
                Next
        End Select
nextsec:
    Loop
    EV_DoFloor = rtn
End Function

' BUILD A STAIRCASE!
Public Function EV_BuildStairs(ByVal line As Long, ByVal typ As Long) As Long
    Dim secnum As Long, height As Long, i As Long, newsecnum As Long, texture As Long
    Dim ok As Boolean, rtn As Long, sec As Long, tsec As Long, floor As Long
    Dim stairsize As Long, speed As Long, ln As Long

    secnum = -1
    rtn = 0
    Do
        secnum = P_FindSectorFromLineTag(line, secnum)
        If secnum < 0 Then Exit Do

        sec = secnum

        ' ALREADY MOVING?  IF SO, KEEP GOING...
        If sectors(sec).specialdata <> 0 Then GoTo nextsec

        ' new floor thinker
        rtn = 1
        floor = P_NewThinker()
        P_AddThinker floor
        sectors(sec).specialdata = floor
        thinkers(floor).func = TF_FLOOR
        floormoves(floor).direction = 1
        floormoves(floor).sector = sec
        Select Case typ
            Case build8
                speed = FLOORSPEED \ 4
                stairsize = 8 * FRACUNIT
            Case turbo16
                speed = FLOORSPEED * 4
                stairsize = 16 * FRACUNIT
        End Select
        floormoves(floor).speed = speed
        height = sectors(sec).floorheight + stairsize
        floormoves(floor).floordestheight = height

        texture = sectors(sec).floorpic

        ' Find next sector to raise
        ' 1.  Find 2-sided line with same sector side[0]
        ' 2.  Other side is the next sector to raise
        Do
            ok = False
            For i = 0 To sectors(sec).linecount - 1
                ln = linebuffer(sectors(sec).lines + i)
                If (lines(ln).flags And ML_TWOSIDED) = 0 Then GoTo nextline

                tsec = lines(ln).frontsector
                newsecnum = tsec
                If secnum <> newsecnum Then GoTo nextline

                tsec = lines(ln).backsector
                newsecnum = tsec

                If sectors(tsec).floorpic <> texture Then GoTo nextline

                height = height + stairsize

                If sectors(tsec).specialdata <> 0 Then GoTo nextline

                sec = tsec
                secnum = newsecnum
                floor = P_NewThinker()

                P_AddThinker floor

                sectors(sec).specialdata = floor
                thinkers(floor).func = TF_FLOOR
                floormoves(floor).direction = 1
                floormoves(floor).sector = sec
                floormoves(floor).speed = speed
                floormoves(floor).floordestheight = height
                ok = True
                Exit For
nextline:
            Next
        Loop While ok
nextsec:
    Loop
    EV_BuildStairs = rtn
End Function
