Attribute VB_Name = "r_interp"
' r_interp.bas - smooth frames between tics (DoomVB only; linuxdoom draws one
' frame per tic).
'
' The game still runs at 35 tics a second, but the screen is drawn once per
' monitor refresh.  At the start of every tic P_Ticker calls R_InterpSave,
' which remembers where every thing, the view, the floors and ceilings, the
' weapon and the scrolling walls are.  A frame drawn between two tics shows
' them part of the way (fractic) from that old place to the new one, so the
' picture is one tic behind the game.  Only the renderer sees these places:
' R_InterpBegin writes the in-between floor and ceiling heights into
' sectors() and R_InterpEnd puts the real ones back, so the game itself, and
' with it demos and net games, does not change.
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public smoothframes As Long                      ' default.cfg smooth_frames (1 = on)
Public smoothactive As Boolean                   ' smooth frames are used in this run

Private stamp As Long                            ' counts the saves (and resets)
Private saved As Boolean                         ' saved since the level was set up
Private interp As Boolean                        ' this frame is drawn between two tics
Private fractic As Long                          ' fixed_t 0..FRACUNIT: how far from the old place
Private ticstartms As Double                     ' I_GetTimeExact at which the saved tic began

' things, by thinker slot
Private mostamp() As Long                        ' stamp of the save that filled the slot
Private oldmox() As Long
Private oldmoy() As Long
Private oldmoz() As Long
Private oldmoangle() As Long
Private momax As Long                            ' upper bound of the arrays above

' view and weapon of each player
Private oldviewz(0 To MAXPLAYERS - 1) As Long
Private oldpspstate(0 To MAXPLAYERS - 1, 0 To NUMPSPRITES - 1) As Long
Private oldpspsx(0 To MAXPLAYERS - 1, 0 To NUMPSPRITES - 1) As Long
Private oldpspsy(0 To MAXPLAYERS - 1, 0 To NUMPSPRITES - 1) As Long

' floors and ceilings: the sectors that moved in the last tic
Private oldfloor() As Long
Private oldceil() As Long
Private secsize As Long                          ' number of sectors the arrays hold
Private movedsec() As Long
Private realfloor() As Long
Private realceil() As Long
Private nummoved As Long
Private movedstamp As Long                       ' stamp for which movedsec() was built

' scrolling walls (line special 48)
Private scrollside() As Long
Private oldscroll() As Long
Private realscroll() As Long
Private numscroll As Long

Private begun As Boolean                         ' R_InterpBegin changed sectors() and sides()

' P_SetupLevel (and so also a loaded savegame): nothing to come from yet.
Public Sub R_InterpReset()
    Dim i As Long
    stamp = stamp + 1
    saved = False
    interp = False
    numscroll = 0
    For i = 0 To numlines - 1
        If lines(i).special = 48 Then
            If numscroll = 0 Then
                ReDim scrollside(0 To 15)
            ElseIf numscroll > UBound(scrollside) Then
                ReDim Preserve scrollside(0 To numscroll * 2)
            End If
            scrollside(numscroll) = lines(i).sidenum(0)
            numscroll = numscroll + 1
        End If
    Next
    If numscroll > 0 Then
        ReDim oldscroll(0 To numscroll - 1)
        ReDim realscroll(0 To numscroll - 1)
    End If
End Sub

' P_NewThinker (a new thing) and P_TeleportMove: draw the thing where it is,
' not on its way from wherever the slot was before.
Public Sub R_InterpForget(ByVal t As Long)
    If t <= momax Then mostamp(t) = 0
End Sub

' P_Ticker, at the start of every tic: remember where everything is.
Public Sub R_InterpSave()
    Dim t As Long, p As Long, i As Long, period As Double

    If Not smoothactive Then Exit Sub
    If (gametic Mod ticdup) <> 0 Then Exit Sub   ' -dup: one step spans the whole group
    stamp = stamp + 1
    saved = True

    ' the tic began at the last tic boundary of the clock (I_GetTime)
    period = 1000# * ticdup / TICRATE
    ticstartms = Int(I_GetTimeExact() / period) * period

    If momax < maxthinkers Then
        momax = maxthinkers
        ReDim Preserve mostamp(0 To momax)
        ReDim Preserve oldmox(0 To momax)
        ReDim Preserve oldmoy(0 To momax)
        ReDim Preserve oldmoz(0 To momax)
        ReDim Preserve oldmoangle(0 To momax)
    End If
    t = thinkers(0).nxt
    Do While t <> 0
        If thinkers(t).func = TF_MOBJ Then
            mostamp(t) = stamp
            oldmox(t) = mobjs(t).x
            oldmoy(t) = mobjs(t).y
            oldmoz(t) = mobjs(t).z
            oldmoangle(t) = mobjs(t).angle
        End If
        t = thinkers(t).nxt
    Loop

    For p = 0 To MAXPLAYERS - 1
        oldviewz(p) = players(p).viewz
        For i = 0 To NUMPSPRITES - 1
            oldpspstate(p, i) = players(p).psprites(i).state
            oldpspsx(p, i) = players(p).psprites(i).sx
            oldpspsy(p, i) = players(p).psprites(i).sy
        Next
    Next

    If secsize <> numsectors Then
        secsize = numsectors
        ReDim oldfloor(0 To secsize - 1)
        ReDim oldceil(0 To secsize - 1)
        ReDim movedsec(0 To secsize - 1)
        ReDim realfloor(0 To secsize - 1)
        ReDim realceil(0 To secsize - 1)
    End If
    For i = 0 To numsectors - 1
        oldfloor(i) = sectors(i).floorheight
        oldceil(i) = sectors(i).ceilingheight
    Next

    For i = 0 To numscroll - 1
        oldscroll(i) = sides(scrollside(i)).textureoffset
    Next
End Sub

' D_DoomLoop, before every frame: how far between the last two tics is the
' moment this frame shows (I_FrameTime, a fixed point in the screen refresh)?
Public Sub R_InterpStartFrame()
    Dim f As Double
    interp = False
    If Not smoothactive Or Not saved Then Exit Sub
    f = (I_FrameTime() - ticstartms) * TICRATE / (1000# * ticdup)
    If f >= 1 Then Exit Sub                      ' the next tic is late: show this one
    If f < 0 Then f = 0
    fractic = Int(f * FRACUNIT)
    interp = True
End Sub

Private Function Lerp(ByVal oldv As Long, ByVal newv As Long) As Long
    Lerp = oldv + FixedMul(newv - oldv, fractic)
End Function

' R_SetupFrame: the view between the last two tics.
Public Sub R_InterpView(ByVal player As Long)
    Dim mo As Long
    If Not interp Then Exit Sub
    mo = players(player).mo
    If mo > momax Then Exit Sub
    If mostamp(mo) <> stamp Then Exit Sub        ' spawned or teleported in the last tic
    viewx = Lerp(oldmox(mo), mobjs(mo).x)
    viewy = Lerp(oldmoy(mo), mobjs(mo).y)
    viewz = Lerp(oldviewz(player), players(player).viewz)
    viewangle = Lerp(oldmoangle(mo), mobjs(mo).angle) + viewangleoffset
End Sub

' R_ProjectSprite: where to draw a thing.
Public Sub R_InterpThing(ByVal thing As Long, x As Long, y As Long, z As Long)
    x = mobjs(thing).x
    y = mobjs(thing).y
    z = mobjs(thing).z
    If Not interp Then Exit Sub
    If thing > momax Then Exit Sub
    If mostamp(thing) <> stamp Then Exit Sub
    x = Lerp(oldmox(thing), x)
    y = Lerp(oldmoy(thing), y)
    z = Lerp(oldmoz(thing), z)
End Sub

' R_DrawPSprite: where to draw the weapon.  Only while it shows the same
' frame, so the bob and raising / lowering glide but firing does not.
Public Sub R_InterpPSprite(ByVal player As Long, ByVal pspnum As Long, sx As Long, sy As Long)
    Dim st As Long, ost As Long
    sx = players(player).psprites(pspnum).sx
    sy = players(player).psprites(pspnum).sy
    If Not interp Then Exit Sub
    st = players(player).psprites(pspnum).state
    ost = oldpspstate(player, pspnum)
    If ost = S_NULL Then Exit Sub
    If states(st).sprite <> states(ost).sprite Or states(st).frame <> states(ost).frame Then Exit Sub
    sx = Lerp(oldpspsx(player, pspnum), sx)
    sy = Lerp(oldpspsy(player, pspnum), sy)
End Sub

' R_RenderPlayerView, before drawing: moving floors, ceilings and scrolling
' walls go to their in-between places.
Public Sub R_InterpBegin()
    Dim i As Long, s As Long
    If Not interp Then Exit Sub
    If secsize <> numsectors Then Exit Sub

    ' which sectors moved in the last tic (once per tic)
    If movedstamp <> stamp Then
        nummoved = 0
        For i = 0 To numsectors - 1
            If sectors(i).floorheight <> oldfloor(i) Or sectors(i).ceilingheight <> oldceil(i) Then
                movedsec(nummoved) = i
                nummoved = nummoved + 1
            End If
        Next
        movedstamp = stamp
    End If

    For i = 0 To nummoved - 1
        s = movedsec(i)
        realfloor(i) = sectors(s).floorheight
        realceil(i) = sectors(s).ceilingheight
        sectors(s).floorheight = Lerp(oldfloor(s), realfloor(i))
        sectors(s).ceilingheight = Lerp(oldceil(s), realceil(i))
    Next
    For i = 0 To numscroll - 1
        realscroll(i) = sides(scrollside(i)).textureoffset
        sides(scrollside(i)).textureoffset = Lerp(oldscroll(i), realscroll(i))
    Next
    begun = True
End Sub

' R_RenderPlayerView, after drawing: the game's own heights and offsets.
Public Sub R_InterpEnd()
    Dim i As Long
    If Not begun Then Exit Sub
    For i = 0 To nummoved - 1
        sectors(movedsec(i)).floorheight = realfloor(i)
        sectors(movedsec(i)).ceilingheight = realceil(i)
    Next
    For i = 0 To numscroll - 1
        sides(scrollside(i)).textureoffset = realscroll(i)
    Next
    begun = False
End Sub
