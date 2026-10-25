Attribute VB_Name = "p_lights"
' p_lights.bas - handle sector based lighting updates (p_lights.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' FIRELIGHT FLICKER

' T_FireFlicker
Public Sub T_FireFlicker(ByVal flick As Long)
    Dim amount As Long, sec As Long

    fireflickers(flick).count = fireflickers(flick).count - 1
    If fireflickers(flick).count <> 0 Then Exit Sub

    sec = fireflickers(flick).sector
    amount = (P_Random() And 3) * 16

    If sectors(sec).lightlevel - amount < fireflickers(flick).minlight Then
        sectors(sec).lightlevel = fireflickers(flick).minlight
    Else
        sectors(sec).lightlevel = fireflickers(flick).maxlight - amount
    End If

    fireflickers(flick).count = 4
End Sub

' P_SpawnFireFlicker
Public Sub P_SpawnFireFlicker(ByVal sector As Long)
    Dim flick As Long

    ' Note that we are resetting sector attributes.
    ' Nothing special about it during gameplay.
    sectors(sector).special = 0

    flick = P_NewThinker()
    P_AddThinker flick

    thinkers(flick).func = TF_FIREFLICKER
    fireflickers(flick).sector = sector
    fireflickers(flick).maxlight = sectors(sector).lightlevel
    fireflickers(flick).minlight = P_FindMinSurroundingLight(sector, sectors(sector).lightlevel) + 16
    fireflickers(flick).count = 4
End Sub

' BROKEN LIGHT FLASHING

' T_LightFlash - do flashing lights.
Public Sub T_LightFlash(ByVal flash As Long)
    Dim sec As Long

    lightflashes(flash).count = lightflashes(flash).count - 1
    If lightflashes(flash).count <> 0 Then Exit Sub

    sec = lightflashes(flash).sector
    If sectors(sec).lightlevel = lightflashes(flash).maxlight Then
        sectors(sec).lightlevel = lightflashes(flash).minlight
        lightflashes(flash).count = (P_Random() And lightflashes(flash).mintime) + 1
    Else
        sectors(sec).lightlevel = lightflashes(flash).maxlight
        lightflashes(flash).count = (P_Random() And lightflashes(flash).maxtime) + 1
    End If
End Sub

' P_SpawnLightFlash - after the map has been loaded, scan each sector for
' specials that spawn thinkers
Public Sub P_SpawnLightFlash(ByVal sector As Long)
    Dim flash As Long

    ' nothing special about it during gameplay
    sectors(sector).special = 0

    flash = P_NewThinker()
    P_AddThinker flash

    thinkers(flash).func = TF_LIGHTFLASH
    lightflashes(flash).sector = sector
    lightflashes(flash).maxlight = sectors(sector).lightlevel

    lightflashes(flash).minlight = P_FindMinSurroundingLight(sector, sectors(sector).lightlevel)
    lightflashes(flash).maxtime = 64
    lightflashes(flash).mintime = 7
    lightflashes(flash).count = (P_Random() And lightflashes(flash).maxtime) + 1
End Sub

' STROBE LIGHT FLASHING

' T_StrobeFlash
Public Sub T_StrobeFlash(ByVal flash As Long)
    Dim sec As Long

    strobes(flash).count = strobes(flash).count - 1
    If strobes(flash).count <> 0 Then Exit Sub

    sec = strobes(flash).sector
    If sectors(sec).lightlevel = strobes(flash).minlight Then
        sectors(sec).lightlevel = strobes(flash).maxlight
        strobes(flash).count = strobes(flash).brighttime
    Else
        sectors(sec).lightlevel = strobes(flash).minlight
        strobes(flash).count = strobes(flash).darktime
    End If
End Sub

' P_SpawnStrobeFlash - after the map has been loaded, scan each sector for
' specials that spawn thinkers
Public Sub P_SpawnStrobeFlash(ByVal sector As Long, ByVal fastOrSlow As Long, ByVal inSync As Long)
    Dim flash As Long

    flash = P_NewThinker()
    P_AddThinker flash

    strobes(flash).sector = sector
    strobes(flash).darktime = fastOrSlow
    strobes(flash).brighttime = STROBEBRIGHT
    thinkers(flash).func = TF_STROBE
    strobes(flash).maxlight = sectors(sector).lightlevel
    strobes(flash).minlight = P_FindMinSurroundingLight(sector, sectors(sector).lightlevel)

    If strobes(flash).minlight = strobes(flash).maxlight Then strobes(flash).minlight = 0

    ' nothing special about it during gameplay
    sectors(sector).special = 0

    If inSync = 0 Then
        strobes(flash).count = (P_Random() And 7) + 1
    Else
        strobes(flash).count = 1
    End If
End Sub

' Start strobing lights (usually from a trigger)
Public Sub EV_StartLightStrobing(ByVal line As Long)
    Dim secnum As Long

    secnum = -1
    Do
        secnum = P_FindSectorFromLineTag(line, secnum)
        If secnum < 0 Then Exit Do
        If sectors(secnum).specialdata = 0 Then P_SpawnStrobeFlash secnum, SLOWDARK, 0
    Loop
End Sub

' TURN LINE'S TAG LIGHTS OFF
Public Sub EV_TurnTagLightsOff(ByVal line As Long)
    Dim i As Long, j As Long, min As Long, tsec As Long

    For j = 0 To numsectors - 1
        If sectors(j).tag = lines(line).tag Then
            min = sectors(j).lightlevel
            For i = 0 To sectors(j).linecount - 1
                tsec = getNextSector(linebuffer(sectors(j).lines + i), j)
                If tsec <> -1 Then
                    If sectors(tsec).lightlevel < min Then min = sectors(tsec).lightlevel
                End If
            Next
            sectors(j).lightlevel = min
        End If
    Next
End Sub

' TURN LINE'S TAG LIGHTS ON
Public Sub EV_LightTurnOn(ByVal line As Long, ByVal bright As Long)
    Dim i As Long, j As Long, temp As Long

    For i = 0 To numsectors - 1
        If sectors(i).tag = lines(line).tag Then
            ' bright = 0 means to search for highest light level surrounding sector
            If bright = 0 Then
                For j = 0 To sectors(i).linecount - 1
                    temp = getNextSector(linebuffer(sectors(i).lines + j), i)
                    If temp <> -1 Then
                        If sectors(temp).lightlevel > bright Then bright = sectors(temp).lightlevel
                    End If
                Next
            End If
            sectors(i).lightlevel = bright
        End If
    Next
End Sub

' Spawn glowing light
Public Sub T_Glow(ByVal g As Long)
    Dim sec As Long
    sec = glows(g).sector
    Select Case glows(g).direction
        Case -1
            ' DOWN
            sectors(sec).lightlevel = sectors(sec).lightlevel - GLOWSPEED
            If sectors(sec).lightlevel <= glows(g).minlight Then
                sectors(sec).lightlevel = sectors(sec).lightlevel + GLOWSPEED
                glows(g).direction = 1
            End If

        Case 1
            ' UP
            sectors(sec).lightlevel = sectors(sec).lightlevel + GLOWSPEED
            If sectors(sec).lightlevel >= glows(g).maxlight Then
                sectors(sec).lightlevel = sectors(sec).lightlevel - GLOWSPEED
                glows(g).direction = -1
            End If
    End Select
End Sub

Public Sub P_SpawnGlowingLight(ByVal sector As Long)
    Dim g As Long

    g = P_NewThinker()
    P_AddThinker g

    glows(g).sector = sector
    glows(g).minlight = P_FindMinSurroundingLight(sector, sectors(sector).lightlevel)
    glows(g).maxlight = sectors(sector).lightlevel
    thinkers(g).func = TF_GLOW
    glows(g).direction = -1

    sectors(sector).special = 0
End Sub
