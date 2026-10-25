Attribute VB_Name = "p_doors"
' p_doors.bas - door animation code (opening/closing) (p_doors.c).
' (The unused sliding door code of the C file is not ported.)
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' VERTICAL DOORS

' T_VerticalDoor
Public Sub T_VerticalDoor(ByVal door As Long)
    Dim res As Long, sec As Long

    sec = doors(door).sector

    Select Case doors(door).direction
        Case 0
            ' WAITING
            doors(door).topcountdown = doors(door).topcountdown - 1
            If doors(door).topcountdown = 0 Then
                Select Case doors(door).typ
                    Case vld_blazeRaise
                        doors(door).direction = -1 ' time to go back down
                        S_StartSound SectorSoundOrg(sec), sfx_bdcls
                    Case vld_normal
                        doors(door).direction = -1 ' time to go back down
                        S_StartSound SectorSoundOrg(sec), sfx_dorcls
                    Case vld_close30ThenOpen
                        doors(door).direction = 1
                        S_StartSound SectorSoundOrg(sec), sfx_doropn
                End Select
            End If

        Case 2
            '  INITIAL WAIT
            doors(door).topcountdown = doors(door).topcountdown - 1
            If doors(door).topcountdown = 0 Then
                Select Case doors(door).typ
                    Case vld_raiseIn5Mins
                        doors(door).direction = 1
                        doors(door).typ = vld_normal
                        S_StartSound SectorSoundOrg(sec), sfx_doropn
                End Select
            End If

        Case -1
            ' DOWN
            res = T_MovePlane(sec, doors(door).speed, sectors(sec).floorheight, False, 1, doors(door).direction)
            If res = res_pastdest Then
                Select Case doors(door).typ
                    Case vld_blazeRaise, vld_blazeClose
                        sectors(sec).specialdata = 0
                        P_RemoveThinker door      ' unlink and free
                        S_StartSound SectorSoundOrg(sec), sfx_bdcls
                    Case vld_normal, vld_close
                        sectors(sec).specialdata = 0
                        P_RemoveThinker door      ' unlink and free
                    Case vld_close30ThenOpen
                        doors(door).direction = 0
                        doors(door).topcountdown = 35 * 30
                End Select
            ElseIf res = res_crushed Then
                Select Case doors(door).typ
                    Case vld_blazeClose, vld_close ' DO NOT GO BACK UP!
                    Case Else
                        doors(door).direction = 1
                        S_StartSound SectorSoundOrg(sec), sfx_doropn
                End Select
            End If

        Case 1
            ' UP
            res = T_MovePlane(sec, doors(door).speed, doors(door).topheight, False, 1, doors(door).direction)

            If res = res_pastdest Then
                Select Case doors(door).typ
                    Case vld_blazeRaise, vld_normal
                        doors(door).direction = 0 ' wait at top
                        doors(door).topcountdown = doors(door).topwait
                    Case vld_close30ThenOpen, vld_blazeOpen, vld_open
                        sectors(sec).specialdata = 0
                        P_RemoveThinker door      ' unlink and free
                End Select
            End If
    End Select
End Sub

' EV_DoLockedDoor - move a locked door up/down
Public Function EV_DoLockedDoor(ByVal line As Long, ByVal typ As Long, ByVal thing As Long) As Long
    Dim p As Long

    p = mobjs(thing).player

    If p < 0 Then
        EV_DoLockedDoor = 0
        Exit Function
    End If

    Select Case lines(line).special
        Case 99, 133                             ' Blue Lock
            If Not players(p).cards(it_bluecard) And Not players(p).cards(it_blueskull) Then
                players(p).message = PD_BLUEO
                S_StartSound 0, sfx_oof
                EV_DoLockedDoor = 0
                Exit Function
            End If

        Case 134, 135                            ' Red Lock
            If Not players(p).cards(it_redcard) And Not players(p).cards(it_redskull) Then
                players(p).message = PD_REDO
                S_StartSound 0, sfx_oof
                EV_DoLockedDoor = 0
                Exit Function
            End If

        Case 136, 137                            ' Yellow Lock
            If Not players(p).cards(it_yellowcard) And Not players(p).cards(it_yellowskull) Then
                players(p).message = PD_YELLOWO
                S_StartSound 0, sfx_oof
                EV_DoLockedDoor = 0
                Exit Function
            End If
    End Select

    EV_DoLockedDoor = EV_DoDoor(line, typ)
End Function

Public Function EV_DoDoor(ByVal line As Long, ByVal typ As Long) As Long
    Dim secnum As Long, rtn As Long, door As Long

    secnum = -1
    rtn = 0

    Do
        secnum = P_FindSectorFromLineTag(line, secnum)
        If secnum < 0 Then Exit Do

        If sectors(secnum).specialdata <> 0 Then GoTo nextsec

        ' new door thinker
        rtn = 1
        door = P_NewThinker()
        P_AddThinker door
        sectors(secnum).specialdata = door

        thinkers(door).func = TF_DOOR
        doors(door).sector = secnum
        doors(door).typ = typ
        doors(door).topwait = VDOORWAIT
        doors(door).speed = VDOORSPEED

        Select Case typ
            Case vld_blazeClose
                doors(door).topheight = P_FindLowestCeilingSurrounding(secnum)
                doors(door).topheight = doors(door).topheight - 4 * FRACUNIT
                doors(door).direction = -1
                doors(door).speed = VDOORSPEED * 4
                S_StartSound SectorSoundOrg(secnum), sfx_bdcls

            Case vld_close
                doors(door).topheight = P_FindLowestCeilingSurrounding(secnum)
                doors(door).topheight = doors(door).topheight - 4 * FRACUNIT
                doors(door).direction = -1
                S_StartSound SectorSoundOrg(secnum), sfx_dorcls

            Case vld_close30ThenOpen
                doors(door).topheight = sectors(secnum).ceilingheight
                doors(door).direction = -1
                S_StartSound SectorSoundOrg(secnum), sfx_dorcls

            Case vld_blazeRaise, vld_blazeOpen
                doors(door).direction = 1
                doors(door).topheight = P_FindLowestCeilingSurrounding(secnum)
                doors(door).topheight = doors(door).topheight - 4 * FRACUNIT
                doors(door).speed = VDOORSPEED * 4
                If doors(door).topheight <> sectors(secnum).ceilingheight Then
                    S_StartSound SectorSoundOrg(secnum), sfx_bdopn
                End If

            Case vld_normal, vld_open
                doors(door).direction = 1
                doors(door).topheight = P_FindLowestCeilingSurrounding(secnum)
                doors(door).topheight = doors(door).topheight - 4 * FRACUNIT
                If doors(door).topheight <> sectors(secnum).ceilingheight Then
                    S_StartSound SectorSoundOrg(secnum), sfx_doropn
                End If
        End Select
nextsec:
    Loop
    EV_DoDoor = rtn
End Function

' EV_VerticalDoor : open a door manually, no tag value
Public Sub EV_VerticalDoor(ByVal line As Long, ByVal thing As Long)
    Dim player As Long, sec As Long, door As Long, side As Long

    side = 0                                     ' only front sides can be used

    ' Check for locks
    player = mobjs(thing).player

    Select Case lines(line).special
        Case 26, 32                              ' Blue Lock
            If player < 0 Then Exit Sub
            If Not players(player).cards(it_bluecard) And Not players(player).cards(it_blueskull) Then
                players(player).message = PD_BLUEK
                S_StartSound 0, sfx_oof
                Exit Sub
            End If

        Case 27, 34                              ' Yellow Lock
            If player < 0 Then Exit Sub
            If Not players(player).cards(it_yellowcard) And Not players(player).cards(it_yellowskull) Then
                players(player).message = PD_YELLOWK
                S_StartSound 0, sfx_oof
                Exit Sub
            End If

        Case 28, 33                              ' Red Lock
            If player < 0 Then Exit Sub
            If Not players(player).cards(it_redcard) And Not players(player).cards(it_redskull) Then
                players(player).message = PD_REDK
                S_StartSound 0, sfx_oof
                Exit Sub
            End If
    End Select

    ' if the sector has an active thinker, use it
    sec = sides(lines(line).sidenum(side Xor 1)).sector

    If sectors(sec).specialdata <> 0 Then
        door = sectors(sec).specialdata
        Select Case lines(line).special
            Case 1, 26, 27, 28, 117
                ' ONLY FOR "RAISE" DOORS, NOT "OPEN"s
                If doors(door).direction = -1 Then
                    doors(door).direction = 1    ' go back up
                Else
                    If mobjs(thing).player < 0 Then Exit Sub ' JDC: bad guys never close doors

                    doors(door).direction = -1   ' start going down immediately
                End If
                Exit Sub
        End Select
    End If

    ' for proper sound
    Select Case lines(line).special
        Case 117, 118                            ' BLAZING DOOR RAISE / OPEN
            S_StartSound SectorSoundOrg(sec), sfx_bdopn
        Case 1, 31                               ' NORMAL DOOR SOUND
            S_StartSound SectorSoundOrg(sec), sfx_doropn
        Case Else                                ' LOCKED DOOR SOUND
            S_StartSound SectorSoundOrg(sec), sfx_doropn
    End Select

    ' new door thinker
    door = P_NewThinker()
    P_AddThinker door
    sectors(sec).specialdata = door
    thinkers(door).func = TF_DOOR
    doors(door).sector = sec
    doors(door).direction = 1
    doors(door).speed = VDOORSPEED
    doors(door).topwait = VDOORWAIT

    Select Case lines(line).special
        Case 1, 26, 27, 28
            doors(door).typ = vld_normal
        Case 31, 32, 33, 34
            doors(door).typ = vld_open
            lines(line).special = 0
        Case 117                                 ' blazing door raise
            doors(door).typ = vld_blazeRaise
            doors(door).speed = VDOORSPEED * 4
        Case 118                                 ' blazing door open
            doors(door).typ = vld_blazeOpen
            lines(line).special = 0
            doors(door).speed = VDOORSPEED * 4
    End Select

    ' find the top and bottom of the movement range
    doors(door).topheight = P_FindLowestCeilingSurrounding(sec)
    doors(door).topheight = doors(door).topheight - 4 * FRACUNIT
End Sub

' Spawn a door that closes after 30 seconds
Public Sub P_SpawnDoorCloseIn30(ByVal sec As Long)
    Dim door As Long

    door = P_NewThinker()
    P_AddThinker door

    sectors(sec).specialdata = door
    sectors(sec).special = 0

    thinkers(door).func = TF_DOOR
    doors(door).sector = sec
    doors(door).direction = 0
    doors(door).typ = vld_normal
    doors(door).speed = VDOORSPEED
    doors(door).topcountdown = 30 * 35
End Sub

' Spawn a door that opens after 5 minutes
Public Sub P_SpawnDoorRaiseIn5Mins(ByVal sec As Long, ByVal secnum As Long)
    Dim door As Long

    door = P_NewThinker()
    P_AddThinker door

    sectors(sec).specialdata = door
    sectors(sec).special = 0

    thinkers(door).func = TF_DOOR
    doors(door).sector = sec
    doors(door).direction = 2
    doors(door).typ = vld_raiseIn5Mins
    doors(door).speed = VDOORSPEED
    doors(door).topheight = P_FindLowestCeilingSurrounding(sec)
    doors(door).topheight = doors(door).topheight - 4 * FRACUNIT
    doors(door).topwait = VDOORWAIT
    doors(door).topcountdown = 5 * 60 * 35
End Sub
