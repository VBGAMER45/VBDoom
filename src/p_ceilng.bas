Attribute VB_Name = "p_ceilng"
' p_ceilng.bas - ceiling aninmation (lowering, crushing, raising) (p_ceilng.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' CEILINGS
' (activeceilings() lives in doomdef)

' T_MoveCeiling
Public Sub T_MoveCeiling(ByVal ceiling As Long)
    Dim res As Long, sec As Long

    sec = ceilings(ceiling).sector

    Select Case ceilings(ceiling).direction
        Case 0
            ' IN STASIS

        Case 1
            ' UP
            res = T_MovePlane(sec, ceilings(ceiling).speed, ceilings(ceiling).topheight, False, 1, ceilings(ceiling).direction)

            If (leveltime And 7) = 0 Then
                If ceilings(ceiling).typ <> silentCrushAndRaise Then S_StartSound SectorSoundOrg(sec), sfx_stnmov
            End If

            If res = res_pastdest Then
                Select Case ceilings(ceiling).typ
                    Case raiseToHighest
                        P_RemoveActiveCeiling ceiling
                    Case silentCrushAndRaise
                        S_StartSound SectorSoundOrg(sec), sfx_pstop
                        ceilings(ceiling).direction = -1
                    Case fastCrushAndRaise, crushAndRaise
                        ceilings(ceiling).direction = -1
                End Select
            End If

        Case -1
            ' DOWN
            res = T_MovePlane(sec, ceilings(ceiling).speed, ceilings(ceiling).bottomheight, _
                              ceilings(ceiling).crush, 1, ceilings(ceiling).direction)

            If (leveltime And 7) = 0 Then
                If ceilings(ceiling).typ <> silentCrushAndRaise Then S_StartSound SectorSoundOrg(sec), sfx_stnmov
            End If

            If res = res_pastdest Then
                Select Case ceilings(ceiling).typ
                    Case silentCrushAndRaise
                        S_StartSound SectorSoundOrg(sec), sfx_pstop
                        ceilings(ceiling).speed = CEILSPEED
                        ceilings(ceiling).direction = 1
                    Case crushAndRaise
                        ceilings(ceiling).speed = CEILSPEED
                        ceilings(ceiling).direction = 1
                    Case fastCrushAndRaise
                        ceilings(ceiling).direction = 1
                    Case lowerAndCrush, lowerToFloor
                        P_RemoveActiveCeiling ceiling
                End Select
            ElseIf res = res_crushed Then
                Select Case ceilings(ceiling).typ
                    Case silentCrushAndRaise, crushAndRaise, lowerAndCrush
                        ceilings(ceiling).speed = CEILSPEED \ 8
                End Select
            End If
    End Select
End Sub

' EV_DoCeiling - move a ceiling up/down and all around!
Public Function EV_DoCeiling(ByVal line As Long, ByVal typ As Long) As Long
    Dim secnum As Long, rtn As Long, sec As Long, ceiling As Long

    secnum = -1
    rtn = 0

    ' Reactivate in-stasis ceilings...for certain types.
    Select Case typ
        Case fastCrushAndRaise, silentCrushAndRaise, crushAndRaise
            P_ActivateInStasisCeiling line
    End Select

    Do
        secnum = P_FindSectorFromLineTag(line, secnum)
        If secnum < 0 Then Exit Do

        sec = secnum
        If sectors(sec).specialdata <> 0 Then GoTo nextsec

        ' new door thinker
        rtn = 1
        ceiling = P_NewThinker()
        P_AddThinker ceiling
        sectors(sec).specialdata = ceiling
        thinkers(ceiling).func = TF_CEILING
        ceilings(ceiling).sector = sec
        ceilings(ceiling).crush = False

        Select Case typ
            Case fastCrushAndRaise
                ceilings(ceiling).crush = True
                ceilings(ceiling).topheight = sectors(sec).ceilingheight
                ceilings(ceiling).bottomheight = sectors(sec).floorheight + (8 * FRACUNIT)
                ceilings(ceiling).direction = -1
                ceilings(ceiling).speed = CEILSPEED * 2

            Case silentCrushAndRaise, crushAndRaise, lowerAndCrush, lowerToFloor
                If typ = silentCrushAndRaise Or typ = crushAndRaise Then
                    ceilings(ceiling).crush = True
                    ceilings(ceiling).topheight = sectors(sec).ceilingheight
                End If
                ceilings(ceiling).bottomheight = sectors(sec).floorheight
                If typ <> lowerToFloor Then
                    ceilings(ceiling).bottomheight = ceilings(ceiling).bottomheight + 8 * FRACUNIT
                End If
                ceilings(ceiling).direction = -1
                ceilings(ceiling).speed = CEILSPEED

            Case raiseToHighest
                ceilings(ceiling).topheight = P_FindHighestCeilingSurrounding(sec)
                ceilings(ceiling).direction = 1
                ceilings(ceiling).speed = CEILSPEED
        End Select

        ceilings(ceiling).tag = sectors(sec).tag
        ceilings(ceiling).typ = typ
        P_AddActiveCeiling ceiling
nextsec:
    Loop
    EV_DoCeiling = rtn
End Function

' Add an active ceiling
Public Sub P_AddActiveCeiling(ByVal c As Long)
    Dim i As Long
    For i = 0 To MAXCEILINGS - 1
        If activeceilings(i) = 0 Then
            activeceilings(i) = c
            Exit Sub
        End If
    Next
End Sub

' Remove a ceiling's thinker
Public Sub P_RemoveActiveCeiling(ByVal c As Long)
    Dim i As Long
    For i = 0 To MAXCEILINGS - 1
        If activeceilings(i) = c Then
            sectors(ceilings(c).sector).specialdata = 0
            P_RemoveThinker c
            activeceilings(i) = 0
            Exit For
        End If
    Next
End Sub

' Restart a ceiling that's in-stasis
Public Sub P_ActivateInStasisCeiling(ByVal line As Long)
    Dim i As Long, c As Long
    For i = 0 To MAXCEILINGS - 1
        c = activeceilings(i)
        If c <> 0 Then
            If ceilings(c).tag = lines(line).tag And ceilings(c).direction = 0 Then
                ceilings(c).direction = ceilings(c).olddirection
                thinkers(c).func = TF_CEILING
            End If
        End If
    Next
End Sub

' EV_CeilingCrushStop - stop a ceiling from crushing!
Public Function EV_CeilingCrushStop(ByVal line As Long) As Long
    Dim i As Long, rtn As Long, c As Long

    rtn = 0
    For i = 0 To MAXCEILINGS - 1
        c = activeceilings(i)
        If c <> 0 Then
            If ceilings(c).tag = lines(line).tag And ceilings(c).direction <> 0 Then
                ceilings(c).olddirection = ceilings(c).direction
                thinkers(c).func = TF_NONE
                ceilings(c).direction = 0        ' in-stasis
                rtn = 1
            End If
        End If
    Next

    EV_CeilingCrushStop = rtn
End Function
