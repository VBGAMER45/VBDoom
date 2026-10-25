Attribute VB_Name = "p_plats"
' p_plats.bas - plats (i.e. elevator platforms) code, raising/lowering
' (p_plats.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' (activeplats() lives in doomdef)

' Move a plat up and down
Public Sub T_PlatRaise(ByVal plat As Long)
    Dim res As Long, sec As Long

    sec = plats(plat).sector

    Select Case plats(plat).status
        Case plat_up
            res = T_MovePlane(sec, plats(plat).speed, plats(plat).high, plats(plat).crush, 0, 1)

            If plats(plat).typ = raiseAndChange Or plats(plat).typ = raiseToNearestAndChange Then
                If (leveltime And 7) = 0 Then S_StartSound SectorSoundOrg(sec), sfx_stnmov
            End If

            If res = res_crushed And Not plats(plat).crush Then
                plats(plat).count = plats(plat).wait
                plats(plat).status = plat_down
                S_StartSound SectorSoundOrg(sec), sfx_pstart
            ElseIf res = res_pastdest Then
                plats(plat).count = plats(plat).wait
                plats(plat).status = plat_waiting
                S_StartSound SectorSoundOrg(sec), sfx_pstop

                Select Case plats(plat).typ
                    Case blazeDWUS, downWaitUpStay
                        P_RemoveActivePlat plat
                    Case raiseAndChange, raiseToNearestAndChange
                        P_RemoveActivePlat plat
                End Select
            End If

        Case plat_down
            res = T_MovePlane(sec, plats(plat).speed, plats(plat).low, False, 0, -1)

            If res = res_pastdest Then
                plats(plat).count = plats(plat).wait
                plats(plat).status = plat_waiting
                S_StartSound SectorSoundOrg(sec), sfx_pstop
            End If

        Case plat_waiting
            plats(plat).count = plats(plat).count - 1
            If plats(plat).count = 0 Then
                If sectors(sec).floorheight = plats(plat).low Then
                    plats(plat).status = plat_up
                Else
                    plats(plat).status = plat_down
                End If
                S_StartSound SectorSoundOrg(sec), sfx_pstart
            End If

        Case plat_in_stasis
    End Select
End Sub

' Do Platforms
'  "amount" is only used for SOME platforms.
Public Function EV_DoPlat(ByVal line As Long, ByVal typ As Long, ByVal amount As Long) As Long
    Dim plat As Long, secnum As Long, rtn As Long, sec As Long

    secnum = -1
    rtn = 0

    ' Activate all <type> plats that are in_stasis
    Select Case typ
        Case perpetualRaise
            P_ActivateInStasis lines(line).tag
    End Select

    Do
        secnum = P_FindSectorFromLineTag(line, secnum)
        If secnum < 0 Then Exit Do

        sec = secnum

        If sectors(sec).specialdata <> 0 Then GoTo nextsec

        ' Find lowest & highest floors around sector
        rtn = 1
        plat = P_NewThinker()
        P_AddThinker plat

        plats(plat).typ = typ
        plats(plat).sector = sec
        sectors(sec).specialdata = plat
        thinkers(plat).func = TF_PLAT
        plats(plat).crush = False
        plats(plat).tag = lines(line).tag

        Select Case typ
            Case raiseToNearestAndChange
                plats(plat).speed = PLATSPEED \ 2
                sectors(sec).floorpic = sectors(sides(lines(line).sidenum(0)).sector).floorpic
                plats(plat).high = P_FindNextHighestFloor(sec, sectors(sec).floorheight)
                plats(plat).wait = 0
                plats(plat).status = plat_up
                ' NO MORE DAMAGE, IF APPLICABLE
                sectors(sec).special = 0

                S_StartSound SectorSoundOrg(sec), sfx_stnmov

            Case raiseAndChange
                plats(plat).speed = PLATSPEED \ 2
                sectors(sec).floorpic = sectors(sides(lines(line).sidenum(0)).sector).floorpic
                plats(plat).high = sectors(sec).floorheight + amount * FRACUNIT
                plats(plat).wait = 0
                plats(plat).status = plat_up

                S_StartSound SectorSoundOrg(sec), sfx_stnmov

            Case downWaitUpStay
                plats(plat).speed = PLATSPEED * 4
                plats(plat).low = P_FindLowestFloorSurrounding(sec)

                If plats(plat).low > sectors(sec).floorheight Then plats(plat).low = sectors(sec).floorheight

                plats(plat).high = sectors(sec).floorheight
                plats(plat).wait = 35 * PLATWAIT
                plats(plat).status = plat_down
                S_StartSound SectorSoundOrg(sec), sfx_pstart

            Case blazeDWUS
                plats(plat).speed = PLATSPEED * 8
                plats(plat).low = P_FindLowestFloorSurrounding(sec)

                If plats(plat).low > sectors(sec).floorheight Then plats(plat).low = sectors(sec).floorheight

                plats(plat).high = sectors(sec).floorheight
                plats(plat).wait = 35 * PLATWAIT
                plats(plat).status = plat_down
                S_StartSound SectorSoundOrg(sec), sfx_pstart

            Case perpetualRaise
                plats(plat).speed = PLATSPEED
                plats(plat).low = P_FindLowestFloorSurrounding(sec)

                If plats(plat).low > sectors(sec).floorheight Then plats(plat).low = sectors(sec).floorheight

                plats(plat).high = P_FindHighestFloorSurrounding(sec)

                If plats(plat).high < sectors(sec).floorheight Then plats(plat).high = sectors(sec).floorheight

                plats(plat).wait = 35 * PLATWAIT
                plats(plat).status = P_Random() And 1

                S_StartSound SectorSoundOrg(sec), sfx_pstart
        End Select
        P_AddActivePlat plat
nextsec:
    Loop
    EV_DoPlat = rtn
End Function

Public Sub P_ActivateInStasis(ByVal tag As Long)
    Dim i As Long, p As Long
    For i = 0 To MAXPLATS - 1
        p = activeplats(i)
        If p <> 0 Then
            If plats(p).tag = tag And plats(p).status = plat_in_stasis Then
                plats(p).status = plats(p).oldstatus
                thinkers(p).func = TF_PLAT
            End If
        End If
    Next
End Sub

Public Sub EV_StopPlat(ByVal line As Long)
    Dim j As Long, p As Long
    For j = 0 To MAXPLATS - 1
        p = activeplats(j)
        If p <> 0 Then
            If plats(p).status <> plat_in_stasis And plats(p).tag = lines(line).tag Then
                plats(p).oldstatus = plats(p).status
                plats(p).status = plat_in_stasis
                thinkers(p).func = TF_NONE
            End If
        End If
    Next
End Sub

Public Sub P_AddActivePlat(ByVal plat As Long)
    Dim i As Long
    For i = 0 To MAXPLATS - 1
        If activeplats(i) = 0 Then
            activeplats(i) = plat
            Exit Sub
        End If
    Next
    I_Error "P_AddActivePlat: no more plats!"
End Sub

Public Sub P_RemoveActivePlat(ByVal plat As Long)
    Dim i As Long
    For i = 0 To MAXPLATS - 1
        If plat = activeplats(i) Then
            sectors(plats(plat).sector).specialdata = 0
            P_RemoveThinker plat
            activeplats(i) = 0
            Exit Sub
        End If
    Next
    I_Error "P_RemoveActivePlat: can't find plat!"
End Sub
