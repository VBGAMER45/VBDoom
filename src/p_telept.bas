Attribute VB_Name = "p_telept"
' p_telept.bas - teleportation (p_telept.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' TELEPORTATION
Public Function EV_Teleport(ByVal line As Long, ByVal side As Long, ByVal thing As Long) As Long
    Dim i As Long, tag As Long, m As Long, th As Long, mover As Long

    ' don't teleport missiles
    If (mobjs(thing).flags And MF_MISSILE) <> 0 Then
        EV_Teleport = 0
        Exit Function
    End If

    ' Don't teleport if hit back of line, so you can get out of teleporter.
    If side = 1 Then
        EV_Teleport = 0
        Exit Function
    End If

    tag = lines(line).tag
    mover = thing
    For i = 0 To numsectors - 1
        If sectors(i).tag = tag Then
            th = thinkers(0).nxt
            Do While th <> 0
                m = th
                th = thinkers(th).nxt
                ' a mobj that is a teleportman in the tagged sector?
                If thinkers(m).func = TF_MOBJ Then
                    If mobjs(m).mtype = MT_TELEPORTMAN Then
                        If subsectors(mobjs(m).subsector).sector = i Then
                            EV_Teleport = TeleportTo(mover, m)
                            Exit Function
                        End If
                    End If
                End If
            Loop
        End If
    Next
    EV_Teleport = 0
End Function

' moves "thing" to the teleport destination "m" (the body of the C loop)
Private Function TeleportTo(ByVal thing As Long, ByVal m As Long) As Long
    Dim fog As Long, an As Long, oldx As Long, oldy As Long, oldz As Long, p As Long
    Dim destx As Long, desty As Long

    oldx = mobjs(thing).x
    oldy = mobjs(thing).y
    oldz = mobjs(thing).z

    destx = mobjs(m).x
    desty = mobjs(m).y
    If Not P_TeleportMove(thing, destx, desty) Then
        TeleportTo = 0
        Exit Function
    End If

    mobjs(thing).z = mobjs(thing).floorz       ' fixme: not needed?
    p = mobjs(thing).player
    If p >= 0 Then players(p).viewz = mobjs(thing).z + players(p).viewheight

    ' spawn teleport fog at source and destination
    fog = P_SpawnMobj(oldx, oldy, oldz, MT_TFOG)
    S_StartSound fog, sfx_telept
    an = FINEANGLE(mobjs(m).angle)
    fog = P_SpawnMobj(mobjs(m).x + 20 * finecosine(an), mobjs(m).y + 20 * finesine(an), _
                      mobjs(thing).z, MT_TFOG)

    ' emit sound, where?
    S_StartSound fog, sfx_telept

    ' don't move for a bit
    If p >= 0 Then mobjs(thing).reactiontime = 18

    mobjs(thing).angle = mobjs(m).angle
    mobjs(thing).momx = 0
    mobjs(thing).momy = 0
    mobjs(thing).momz = 0
    TeleportTo = 1
End Function
