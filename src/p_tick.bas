Attribute VB_Name = "p_tick"
' p_tick.bas - archiving: thinker list handling and the play loop ticker
' (p_tick.c).
'
' Every thinker (mobj, door, platform, light, ...) occupies one slot of the
' shared arrays declared in doomdef: thinkers(slot) holds the list links and
' the function id, mobjs(slot) / ceilings(slot) / ... the data.  Slot 0 is
' thinkercap.  Removed slots go back on a free list.
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const INITIALTHINKERS As Long = 8192
Private Const FREERESERVE As Long = 2048         ' grow between tics when fewer slots are free

Private freeslots() As Long
Public checklinks As Boolean                     ' -checklinks debug option
Public currentThinker As Long
Public dbgUnsetBy As String                      ' -checklinks: who called P_UnsetThingPosition
Public synclog As Integer                        ' -synclog <file>: per tic state fingerprint
Public watchslot As Long                         ' -watch <slot>: log the life of one thinker slot
Private slotfree() As Boolean
Private numfree As Long

Private zeroMobj As mobj_t
Private zeroCeiling As ceiling_t
Private zeroDoor As vldoor_t
Private zeroFloor As floormove_t
Private zeroPlat As plat_t
Private zeroFlicker As fireflicker_t
Private zeroFlash As lightflash_t
Private zeroStrobe As strobe_t
Private zeroGlow As glow_t

' Grows every thinker array.  Must never run while a With block or a ByRef
' argument refers to one of them (see docs/CONVENTIONS.md).
Private Sub P_GrowThinkers(ByVal newmax As Long)
    Dim i As Long, oldmax As Long
    oldmax = maxthinkers
    If newmax <= oldmax Then Exit Sub
    ReDim Preserve thinkers(0 To newmax)
    ReDim Preserve mobjs(0 To newmax)
    ReDim Preserve ceilings(0 To newmax)
    ReDim Preserve doors(0 To newmax)
    ReDim Preserve floormoves(0 To newmax)
    ReDim Preserve plats(0 To newmax)
    ReDim Preserve fireflickers(0 To newmax)
    ReDim Preserve lightflashes(0 To newmax)
    ReDim Preserve strobes(0 To newmax)
    ReDim Preserve glows(0 To newmax)
    ReDim Preserve freeslots(0 To newmax)
    ' new slots are free; push them so the lowest numbers come out first
    ReDim Preserve slotfree(0 To newmax)
    For i = newmax To oldmax + 1 Step -1
        freeslots(numfree) = i
        numfree = numfree + 1
        slotfree(i) = True
    Next
    maxthinkers = newmax
End Sub

' P_InitThinkers - all thinker slots become free.
Public Sub P_InitThinkers()
    Dim i As Long
    If maxthinkers = 0 Then
        maxthinkers = INITIALTHINKERS
        ReDim thinkers(0 To maxthinkers)
        ReDim mobjs(0 To maxthinkers)
        ReDim ceilings(0 To maxthinkers)
        ReDim doors(0 To maxthinkers)
        ReDim floormoves(0 To maxthinkers)
        ReDim plats(0 To maxthinkers)
        ReDim fireflickers(0 To maxthinkers)
        ReDim lightflashes(0 To maxthinkers)
        ReDim strobes(0 To maxthinkers)
        ReDim glows(0 To maxthinkers)
        ReDim freeslots(0 To maxthinkers)
    End If
    numfree = 0
    ReDim slotfree(0 To maxthinkers)
    For i = maxthinkers To 1 Step -1
        freeslots(numfree) = i
        numfree = numfree + 1
        thinkers(i).func = TF_REMOVED
        slotfree(i) = True
    Next
    thinkers(0).prv = 0
    thinkers(0).nxt = 0
    thinkers(0).func = TF_NONE
End Sub

' P_NewThinker - allocates a slot (C: Z_Malloc of the thinker's struct) with
' all of its records zeroed.  The thinker is not linked yet.
Public Function P_NewThinker() As Long
    Dim t As Long
    If numfree = 0 Then P_GrowThinkers maxthinkers * 2
    numfree = numfree - 1
    t = freeslots(numfree)
    If checklinks Then
        If UBound(slotfree) < maxthinkers Then ReDim Preserve slotfree(0 To maxthinkers)
        If Not slotfree(t) Then I_Error "P_NewThinker: slot " & t & " handed out but not free (func " & thinkers(t).func & ", type " & mobjs(t).mtype & ") at gametic " & gametic
        slotfree(t) = False
    End If
    thinkers(t).prv = 0
    thinkers(t).nxt = 0
    thinkers(t).func = TF_NONE
    mobjs(t) = zeroMobj
    ceilings(t) = zeroCeiling
    doors(t) = zeroDoor
    floormoves(t) = zeroFloor
    plats(t) = zeroPlat
    fireflickers(t) = zeroFlicker
    lightflashes(t) = zeroFlash
    strobes(t) = zeroStrobe
    glows(t) = zeroGlow
    P_NewThinker = t
End Function

' Savegames: make the thinker arrays exactly n slots big (contents lost).
Public Sub P_SetThinkerCapacity(ByVal n As Long)
    maxthinkers = n
    ReDim thinkers(0 To n)
    ReDim mobjs(0 To n)
    ReDim ceilings(0 To n)
    ReDim doors(0 To n)
    ReDim floormoves(0 To n)
    ReDim plats(0 To n)
    ReDim fireflickers(0 To n)
    ReDim lightflashes(0 To n)
    ReDim strobes(0 To n)
    ReDim glows(0 To n)
    ReDim freeslots(0 To n)
End Sub

' Savegames: after the thinker arrays were restored, every slot that is not
' on the thinker list becomes free again.
Public Sub P_RebuildFreeList()
    Dim inuse() As Boolean, t As Long, i As Long
    ReDim inuse(0 To maxthinkers)
    t = thinkers(0).nxt
    Do While t <> 0
        inuse(t) = True
        t = thinkers(t).nxt
    Loop
    numfree = 0
    For i = maxthinkers To 1 Step -1
        If Not inuse(i) Then
            thinkers(i).func = TF_REMOVED
            freeslots(numfree) = i
            numfree = numfree + 1
        End If
    Next
End Sub

' P_AddThinker - adds a new thinker at the end of the list.
Public Sub P_AddThinker(ByVal th As Long)
    thinkers(thinkers(0).prv).nxt = th
    thinkers(th).nxt = 0
    thinkers(th).prv = thinkers(0).prv
    thinkers(0).prv = th
End Sub

' P_RemoveThinker - deallocation is lazy -- it will not actually be freed
' until its thinking turn comes up.
Public Sub P_RemoveThinker(ByVal th As Long)
    Dim t As Long
    If checklinks Then
        If thinkers(th).func = TF_MOBJ Then
            ' a mobj must already be unlinked from its sector list
            t = sectors(subsectors(mobjs(th).subsector).sector).thinglist
            Do While t <> 0
                If t = th Then
                    I_Error "P_RemoveThinker: mobj " & th & " (type " & mobjs(th).mtype & ") removed while linked; " & _
                            "running thinker " & currentThinker & " func " & thinkers(currentThinker).func & _
                            " type " & mobjs(currentThinker).mtype & " at gametic " & gametic
                End If
                t = mobjs(t).snext
            Loop
        ElseIf thinkers(th).func = TF_REMOVED Then
            I_Error "P_RemoveThinker: slot " & th & " removed twice; running thinker " & currentThinker & _
                    " func " & thinkers(currentThinker).func & " at gametic " & gametic
        End If
    End If
    thinkers(th).func = TF_REMOVED
End Sub

' P_RunThinkers
Private Sub P_RunThinkers()
    Dim cur As Long, nxt As Long

    cur = thinkers(0).nxt
    Do While cur <> 0
        If thinkers(cur).func = TF_REMOVED Then
            ' time to remove it
            thinkers(thinkers(cur).nxt).prv = thinkers(cur).prv
            thinkers(thinkers(cur).prv).nxt = thinkers(cur).nxt
            If checklinks Then
                If slotfree(cur) Then I_Error "P_RunThinkers: slot " & cur & " freed twice (type " & mobjs(cur).mtype & ") at gametic " & gametic
                slotfree(cur) = True
            End If
            freeslots(numfree) = cur
            numfree = numfree + 1
        Else
            currentThinker = cur
            Select Case thinkers(cur).func
                Case TF_MOBJ: P_MobjThinker cur
                Case TF_CEILING: T_MoveCeiling cur
                Case TF_DOOR: T_VerticalDoor cur
                Case TF_FLOOR: T_MoveFloor cur
                Case TF_PLAT: T_PlatRaise cur
                Case TF_FIREFLICKER: T_FireFlicker cur
                Case TF_LIGHTFLASH: T_LightFlash cur
                Case TF_STROBE: T_StrobeFlash cur
                Case TF_GLOW: T_Glow cur
            End Select
        End If
        cur = thinkers(cur).nxt
    Loop
End Sub

' Debug aid (-checklinks): verifies the sector and blockmap thing lists.
Public Sub P_CheckLinks(ByVal where As String)
    Dim sec As Long, t As Long, prv As Long, n As Long, b As Long

    For sec = 0 To numsectors - 1
        t = sectors(sec).thinglist
        prv = 0
        n = 0
        Do While t <> 0
            n = n + 1
            If n > maxthinkers Then I_Error where & ": loop in thing list of sector " & sec & " at gametic " & gametic
            If thinkers(t).func <> TF_MOBJ Then
                I_Error where & ": sector " & sec & " lists dead slot " & t & " (func " & thinkers(t).func & _
                        ", type " & mobjs(t).mtype & ") at gametic " & gametic
            End If
            If mobjs(t).sprev <> prv Then
                I_Error where & ": bad sprev of mobj " & t & " (type " & mobjs(t).mtype & ") in sector " & sec & " at gametic " & gametic
            End If
            If subsectors(mobjs(t).subsector).sector <> sec Then
                I_Error where & ": mobj " & t & " (type " & mobjs(t).mtype & ") listed in sector " & sec & _
                        " but is in sector " & subsectors(mobjs(t).subsector).sector & " at gametic " & gametic
            End If
            prv = t
            t = mobjs(t).snext
        Loop
    Next

    For b = 0 To bmapwidth * bmapheight - 1
        t = blocklinks(b)
        prv = 0
        n = 0
        Do While t <> 0
            n = n + 1
            If n > maxthinkers Then I_Error where & ": loop in block list " & b & " at gametic " & gametic
            If thinkers(t).func <> TF_MOBJ Then
                I_Error where & ": block " & b & " lists dead slot " & t & " (type " & mobjs(t).mtype & ") at gametic " & gametic
            End If
            If mobjs(t).bprev <> prv Then
                I_Error where & ": bad bprev of mobj " & t & " (type " & mobjs(t).mtype & ") in block " & b & " at gametic " & gametic
            End If
            prv = t
            t = mobjs(t).bnext
        Loop
    Next
End Sub

' Debug aid (-synclog): one line per tic with the random index and a checksum
' of every mobj, to find where two runs diverge.
Private Sub P_SyncLog()
    Dim t As Long, sum As Double, n As Long
    t = thinkers(0).nxt
    Do While t <> 0
        If thinkers(t).func = TF_MOBJ Then
            sum = sum + CDbl(mobjs(t).x) * 3 + CDbl(mobjs(t).y) * 7 + CDbl(mobjs(t).z) * 11 + mobjs(t).state * 13 + t
            n = n + 1
        End If
        t = thinkers(t).nxt
    Loop
    Print #synclog, gametic; prndindex; n; sum
End Sub

' P_Ticker
Public Sub P_Ticker()
    Dim i As Long

    ' keep a reserve of free slots so nothing grows the arrays mid-tic
    If numfree < FREERESERVE Then P_GrowThinkers maxthinkers * 2

    ' run the tic
    If paused Then Exit Sub

    ' pause if in menu and at least one tic has been run
    If Not netgame And menuactive And Not demoplayback And players(consoleplayer).viewz <> 1 Then
        Exit Sub
    End If

    For i = 0 To MAXPLAYERS - 1
        If playeringame(i) Then P_PlayerThink i
    Next

    P_RunThinkers
    P_UpdateSpecials
    P_RespawnSpecials

    If checklinks Then P_CheckLinks "P_Ticker"
    If synclog <> 0 Then P_SyncLog

    ' for par times
    leveltime = leveltime + 1
End Sub
