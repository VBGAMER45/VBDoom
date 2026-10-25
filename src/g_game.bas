Attribute VB_Name = "g_game"
' g_game.bas - none (game loop control: ticcmds, level changes, demos,
' save games) (g_game.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const SAVESTRINGSIZE As Long = 24
Private Const VERSIONSIZE As Long = 16
Private Const DEMOMARKER As Long = &H80
Private Const TURBOTHRESHOLD As Long = &H32
Private Const SLOWTURNTICS As Long = 6
Private Const NUMKEYS As Long = 256
Private Const BODYQUESIZE As Long = 32

Public starttime As Long                         ' for comparative timing purposes
Public netdemo As Boolean
Public demoname As String

Private demodata() As Byte                       ' C: demobuffer
Private demo_p As Long                           ' index into demodata()
Private demoend As Long

Private consistancy(0 To 3, 0 To 11) As Long

Private forwardmove(0 To 1) As Long              ' {0x19, 0x32}
Private sidemove(0 To 1) As Long                 ' {0x18, 0x28}
Private angleturn(0 To 2) As Long                ' {640, 1280, 320}  + slow turn
Private movesInit As Boolean

Private gamekeydown(0 To 255) As Boolean
Private turnheld As Long                         ' for accelerative turning

Private mousearray(0 To 3) As Boolean            ' C: mousebuttons = &mousearray[1]
Private mousex As Long
Private mousey As Long

Private dclicktime As Long
Private dclickstate As Boolean
Private dclicks As Long
Private dclicktime2 As Long
Private dclickstate2 As Boolean
Private dclicks2 As Long

Private savegameslot As Long
Private savedescription As String

Private bodyque(0 To 31) As Long

Private secretexit As Boolean
Private savename As String

Private d_skill As Long
Private d_episode As Long
Private d_map As Long
Private defdemoname As String

Private pars(0 To 3, 0 To 9) As Long
Private cpars(0 To 31) As Long
Private parsInit As Boolean

Private zeroPlayer As player_t

Private Sub InitMoves()
    If movesInit Then Exit Sub
    forwardmove(0) = &H19
    forwardmove(1) = &H32
    sidemove(0) = &H18
    sidemove(1) = &H28
    angleturn(0) = 640
    angleturn(1) = 1280
    angleturn(2) = 320                           ' for slow turn
    movesInit = True
End Sub

Public Sub G_SetTurbo(ByVal turboscale As Long)
    InitMoves
    forwardmove(0) = forwardmove(0) * turboscale \ 100
    forwardmove(1) = forwardmove(1) * turboscale \ 100
    sidemove(0) = sidemove(0) * turboscale \ 100
    sidemove(1) = sidemove(1) * turboscale \ 100
End Sub

Private Function MouseButton(ByVal b As Long) As Boolean
    If b >= -1 And b <= 2 Then MouseButton = mousearray(b + 1)
End Function

' wraps a value to a C short
Private Function ToShort(ByVal v As Long) As Long
    v = v And &HFFFF&
    If v >= 32768 Then v = v - 65536
    ToShort = v
End Function

' G_BuildTiccmd - builds a ticcmd from all of the available inputs or reads
' it from the demo buffer.  If recording a demo, write it out.
Public Sub G_BuildTiccmd(cmd As ticcmd_t)
    Dim i As Long, strafe As Boolean, bstrafe As Boolean
    Dim speed As Long, tspeed As Long, forward As Long, side As Long
    Dim blankcmd As ticcmd_t

    InitMoves

    cmd = blankcmd                                ' C: memcpy (cmd, I_BaseTiccmd(), ...)

    cmd.consistancy = consistancy(consoleplayer, maketic Mod BACKUPTICS)

    strafe = gamekeydown(key_strafe) Or MouseButton(mousebstrafe)
    If gamekeydown(key_speed) Then speed = 1 Else speed = 0

    forward = 0
    side = 0

    ' use two stage accelerative turning on the keyboard and joystick
    If gamekeydown(key_right) Or gamekeydown(key_left) Then
        turnheld = turnheld + ticdup
    Else
        turnheld = 0
    End If

    If turnheld < SLOWTURNTICS Then
        tspeed = 2                               ' slow turn
    Else
        tspeed = speed
    End If

    ' let movement keys cancel each other out
    If strafe Then
        If gamekeydown(key_right) Then side = side + sidemove(speed)
        If gamekeydown(key_left) Then side = side - sidemove(speed)
    Else
        If gamekeydown(key_right) Then cmd.angleturn = cmd.angleturn - angleturn(tspeed)
        If gamekeydown(key_left) Then cmd.angleturn = cmd.angleturn + angleturn(tspeed)
    End If

    If gamekeydown(key_up) Then forward = forward + forwardmove(speed)
    If gamekeydown(key_down) Then forward = forward - forwardmove(speed)
    If gamekeydown(key_straferight) Then side = side + sidemove(speed)
    If gamekeydown(key_strafeleft) Then side = side - sidemove(speed)

    ' buttons
    cmd.chatchar = HU_dequeueChatChar()

    If gamekeydown(key_fire) Or MouseButton(mousebfire) Then cmd.buttons = cmd.buttons Or BT_ATTACK

    If gamekeydown(key_use) Then
        cmd.buttons = cmd.buttons Or BT_USE
        ' clear double clicks if hit use button
        dclicks = 0
    End If

    ' chainsaw overrides
    For i = 0 To NUMWEAPONS - 2
        If gamekeydown(49 + i) Then              ' '1' + i
            cmd.buttons = cmd.buttons Or BT_CHANGE
            cmd.buttons = cmd.buttons Or (i * 8)
            Exit For
        End If
    Next

    ' mouse
    If MouseButton(mousebforward) Then forward = forward + forwardmove(speed)

    ' forward double click
    If MouseButton(mousebforward) <> dclickstate And dclicktime > 1 Then
        dclickstate = MouseButton(mousebforward)
        If dclickstate Then dclicks = dclicks + 1
        If dclicks = 2 Then
            cmd.buttons = cmd.buttons Or BT_USE
            dclicks = 0
        Else
            dclicktime = 0
        End If
    Else
        dclicktime = dclicktime + ticdup
        If dclicktime > 20 Then
            dclicks = 0
            dclickstate = False
        End If
    End If

    ' strafe double click
    bstrafe = MouseButton(mousebstrafe)
    If bstrafe <> dclickstate2 And dclicktime2 > 1 Then
        dclickstate2 = bstrafe
        If dclickstate2 Then dclicks2 = dclicks2 + 1
        If dclicks2 = 2 Then
            cmd.buttons = cmd.buttons Or BT_USE
            dclicks2 = 0
        Else
            dclicktime2 = 0
        End If
    Else
        dclicktime2 = dclicktime2 + ticdup
        If dclicktime2 > 20 Then
            dclicks2 = 0
            dclickstate2 = False
        End If
    End If

    forward = forward + mousey
    If strafe Then
        side = side + mousex * 2
    Else
        cmd.angleturn = cmd.angleturn - mousex * 8
    End If

    mousex = 0
    mousey = 0

    If forward > forwardmove(1) Then
        forward = forwardmove(1)
    ElseIf forward < -forwardmove(1) Then
        forward = -forwardmove(1)
    End If
    If side > forwardmove(1) Then
        side = forwardmove(1)
    ElseIf side < -forwardmove(1) Then
        side = -forwardmove(1)
    End If

    cmd.forwardmove = cmd.forwardmove + forward
    cmd.sidemove = cmd.sidemove + side
    cmd.angleturn = ToShort(cmd.angleturn)

    ' special buttons
    If sendpause Then
        sendpause = False
        cmd.buttons = BT_SPECIAL Or BTS_PAUSE
    End If

    If sendsave Then
        sendsave = False
        cmd.buttons = BT_SPECIAL Or BTS_SAVEGAME Or (savegameslot * 4)
    End If
End Sub

' G_DoLoadLevel
Private Sub G_DoLoadLevel()
    Dim i As Long, j As Long

    ' Set the sky map.
    ' First thing, we have a dummy sky texture name, a flat.  The data is in
    ' the WAD only because we look for an actual index, instead of simply
    ' setting one.
    skyflatnum = R_FlatNumForName(SKYFLATNAME)

    ' DOOM determines the sky texture to be used depending on the current
    ' episode, and the game version.
    If gamemode = commercial Then
        skytexture = R_TextureNumForName("SKY3")
        If gamemap < 12 Then
            skytexture = R_TextureNumForName("SKY1")
        ElseIf gamemap < 21 Then
            skytexture = R_TextureNumForName("SKY2")
        End If
    End If

    levelstarttic = gametic                      ' for time calculation

    If wipegamestate = GS_LEVEL Then wipegamestate = -1 ' force a wipe

    gamestate = GS_LEVEL

    For i = 0 To MAXPLAYERS - 1
        If playeringame(i) And players(i).playerstate = PST_DEAD Then players(i).playerstate = PST_REBORN
        For j = 0 To MAXPLAYERS - 1
            players(i).frags(j) = 0
        Next
    Next

    P_SetupLevel gameepisode, gamemap, 0, gameskill
    displayplayer = consoleplayer                ' view the guy you are playing
    starttime = I_GetTime()
    gameaction = ga_nothing

    ' clear cmd building stuff
    For i = 0 To NUMKEYS - 1
        gamekeydown(i) = False
    Next
    mousex = 0
    mousey = 0
    sendpause = False
    sendsave = False
    paused = False
    For i = 0 To 3
        mousearray(i) = False
    Next
End Sub

' G_Responder - get info needed to make ticcmd_ts for the players.
Public Function G_Responder(ev As event_t) As Boolean
    ' allow spy mode changes even during the demo
    If gamestate = GS_LEVEL And ev.typ = ev_keydown And ev.data1 = KEY_F12 And (singledemo Or deathmatch = 0) Then
        ' spy mode
        Do
            displayplayer = displayplayer + 1
            If displayplayer = MAXPLAYERS Then displayplayer = 0
        Loop While Not playeringame(displayplayer) And displayplayer <> consoleplayer
        G_Responder = True
        Exit Function
    End If

    ' any other key pops up menu if in demos
    If gameaction = ga_nothing And Not singledemo And (demoplayback Or gamestate = GS_DEMOSCREEN) Then
        If ev.typ = ev_keydown Or (ev.typ = ev_mouse And ev.data1 <> 0) Or (ev.typ = ev_joystick And ev.data1 <> 0) Then
            M_StartControlPanel
            G_Responder = True
            Exit Function
        End If
        G_Responder = False
        Exit Function
    End If

    If gamestate = GS_LEVEL Then
        If HU_Responder(ev) Then
            G_Responder = True                   ' chat ate the event
            Exit Function
        End If
        If ST_Responder(ev) Then
            G_Responder = True                   ' status window ate it
            Exit Function
        End If
        If AM_Responder(ev) Then
            G_Responder = True                   ' automap ate it
            Exit Function
        End If
    End If

    If gamestate = GS_FINALE Then
        If F_Responder(ev) Then
            G_Responder = True                   ' finale ate the event
            Exit Function
        End If
    End If

    Select Case ev.typ
        Case ev_keydown
            If ev.data1 = KEY_PAUSE Then
                sendpause = True
                G_Responder = True
                Exit Function
            End If
            If ev.data1 >= 0 And ev.data1 < NUMKEYS Then gamekeydown(ev.data1) = True
            G_Responder = True                   ' eat key down events

        Case ev_keyup
            If ev.data1 >= 0 And ev.data1 < NUMKEYS Then gamekeydown(ev.data1) = False
            G_Responder = False                  ' always let key up events filter down

        Case ev_mouse
            mousearray(1) = ((ev.data1 And 1) <> 0)
            mousearray(2) = ((ev.data1 And 2) <> 0)
            mousearray(3) = ((ev.data1 And 4) <> 0)
            mousex = ev.data2 * (mouseSensitivity + 5) \ 10
            mousey = ev.data3 * (mouseSensitivity + 5) \ 10
            G_Responder = True                   ' eat events

        Case Else
            G_Responder = False
    End Select
End Function

' G_Ticker - make ticcmd_ts for the players.
Public Sub G_Ticker()
    Dim i As Long, buf As Long

    ' do player reborns if needed
    For i = 0 To MAXPLAYERS - 1
        If playeringame(i) And players(i).playerstate = PST_REBORN Then G_DoReborn i
    Next

    ' do things to change the game state
    Do While gameaction <> ga_nothing
        Select Case gameaction
            Case ga_loadlevel
                G_DoLoadLevel
            Case ga_newgame
                G_DoNewGame
            Case ga_loadgame
                G_DoLoadGame
            Case ga_savegame
                G_DoSaveGame
            Case ga_playdemo
                G_DoPlayDemo
            Case ga_completed
                G_DoCompleted
            Case ga_victory
                F_StartFinale
                gameaction = ga_nothing
            Case ga_worlddone
                G_DoWorldDone
            Case ga_screenshot
                M_ScreenShot
                gameaction = ga_nothing
        End Select
    Loop

    ' get commands, check consistancy, and build new consistancy check
    buf = (gametic \ ticdup) Mod BACKUPTICS

    For i = 0 To MAXPLAYERS - 1
        If playeringame(i) Then
            players(i).cmd = netcmds(i, buf)

            If demoplayback Then G_ReadDemoTiccmd players(i).cmd
            If demorecording Then G_WriteDemoTiccmd players(i).cmd

            ' check for turbo cheats
            If players(i).cmd.forwardmove > TURBOTHRESHOLD And (gametic And 31) = 0 And ((gametic \ 32) And 3) = i Then
                players(consoleplayer).message = player_names(i) & " is turbo!"
            End If

            If netgame And Not netdemo And (gametic Mod ticdup) = 0 Then
                If gametic > BACKUPTICS And consistancy(i, buf) <> players(i).cmd.consistancy Then
                    I_Error "consistency failure (" & players(i).cmd.consistancy & " should be " & _
                        consistancy(i, buf) & ")"
                End If
                ' (C: a short, so the low 16 bits)
                If players(i).mo <> 0 Then
                    consistancy(i, buf) = mobjs(players(i).mo).x And &HFFFF&
                Else
                    consistancy(i, buf) = rndindex
                End If
                If consistancy(i, buf) >= 32768 Then consistancy(i, buf) = consistancy(i, buf) - 65536
            End If
        End If
    Next

    ' check for special buttons
    For i = 0 To MAXPLAYERS - 1
        If playeringame(i) Then
            If (players(i).cmd.buttons And BT_SPECIAL) <> 0 Then
                Select Case players(i).cmd.buttons And BT_SPECIALMASK
                    Case BTS_PAUSE
                        paused = Not paused
                        If paused Then
                            S_PauseSound
                        Else
                            S_ResumeSound
                        End If
                    Case BTS_SAVEGAME
                        If savedescription = "" Then savedescription = "NET GAME"
                        savegameslot = (players(i).cmd.buttons And BTS_SAVEMASK) \ 4
                        gameaction = ga_savegame
                End Select
            End If
        End If
    Next

    ' do main actions
    Select Case gamestate
        Case GS_LEVEL
            dbgPhase = "P_Ticker"
            P_Ticker
            dbgPhase = "ST_Ticker"
            ST_Ticker
            AM_Ticker
            HU_Ticker
        Case GS_INTERMISSION
            WI_Ticker
        Case GS_FINALE
            F_Ticker
        Case GS_DEMOSCREEN
            D_PageTicker
    End Select
End Sub

' PLAYER STRUCTURE FUNCTIONS
' also see P_SpawnPlayer in P_Things

' G_PlayerFinishLevel - can when a player completes a level
Private Sub G_PlayerFinishLevel(ByVal player As Long)
    Dim i As Long
    For i = 0 To NUMPOWERS - 1
        players(player).powers(i) = 0
    Next
    For i = 0 To NUMCARDS - 1
        players(player).cards(i) = False
    Next
    mobjs(players(player).mo).flags = mobjs(players(player).mo).flags And (Not MF_SHADOW) ' cancel invisibility
    players(player).extralight = 0               ' cancel gun flashes
    players(player).fixedcolormap = 0            ' cancel ir gogles
    players(player).damagecount = 0              ' no palette changes
    players(player).bonuscount = 0
End Sub

' G_PlayerReborn - called after a player dies.  Almost everything is
' cleared and initialized.
Public Sub G_PlayerReborn(ByVal player As Long)
    Dim i As Long, frags(0 To 3) As Long, killcount As Long, itemcount As Long, secretcount As Long

    For i = 0 To 3
        frags(i) = players(player).frags(i)
    Next
    killcount = players(player).killcount
    itemcount = players(player).itemcount
    secretcount = players(player).secretcount

    players(player) = zeroPlayer

    For i = 0 To 3
        players(player).frags(i) = frags(i)
    Next
    players(player).killcount = killcount
    players(player).itemcount = itemcount
    players(player).secretcount = secretcount

    players(player).usedown = True               ' don't do anything immediately
    players(player).attackdown = True
    players(player).playerstate = PST_LIVE
    players(player).health = MAXHEALTH
    players(player).readyweapon = wp_pistol
    players(player).pendingweapon = wp_pistol
    players(player).weaponowned(wp_fist) = True
    players(player).weaponowned(wp_pistol) = True
    players(player).ammo(am_clip) = 50

    For i = 0 To NUMAMMO - 1
        players(player).maxammo(i) = maxammo(i)
    Next
End Sub

' G_CheckSpot - returns false if the player cannot be respawned at the given
' mapthing_t spot because something is occupying it
Private Function G_CheckSpot(ByVal playernum As Long, mthing As mapthing_t) As Boolean
    Dim x As Long, y As Long, ss As Long, an As Long, mo As Long, i As Long

    If players(playernum).mo = 0 Then
        ' first spawn of level, before corpses
        For i = 0 To playernum - 1
            If mobjs(players(i).mo).x = mthing.x * FRACUNIT And mobjs(players(i).mo).y = mthing.y * FRACUNIT Then
                G_CheckSpot = False
                Exit Function
            End If
        Next
        G_CheckSpot = True
        Exit Function
    End If

    x = mthing.x * FRACUNIT
    y = mthing.y * FRACUNIT

    If Not P_CheckPosition(players(playernum).mo, x, y) Then
        G_CheckSpot = False
        Exit Function
    End If

    ' flush an old corpse if needed
    If bodyqueslot >= BODYQUESIZE Then P_RemoveMobj bodyque(bodyqueslot Mod BODYQUESIZE)
    bodyque(bodyqueslot Mod BODYQUESIZE) = players(playernum).mo
    bodyqueslot = bodyqueslot + 1

    ' spawn a teleport fog
    ss = R_PointInSubsector(x, y)
    an = FINEANGLE(ANG45 * (mthing.angle \ 45))

    mo = P_SpawnMobj(x + 20 * finecosine(an), y + 20 * finesine(an), sectors(subsectors(ss).sector).floorheight, MT_TFOG)

    If players(consoleplayer).viewz <> 1 Then S_StartSound mo, sfx_telept ' don't start sound on first frame

    G_CheckSpot = True
End Function

' G_DeathMatchSpawnPlayer - spawns a player at one of the random death match
' spots called at level load and each death
Public Sub G_DeathMatchSpawnPlayer(ByVal playernum As Long)
    Dim i As Long, j As Long, selections As Long, spot As mapthing_t

    selections = deathmatch_p
    If selections < 4 Then I_Error "Only " & selections & " deathmatch spots, 4 required"

    For j = 0 To 19
        i = P_Random() Mod selections
        spot = deathmatchstarts(i)
        If G_CheckSpot(playernum, spot) Then
            deathmatchstarts(i).typ = playernum + 1
            spot = deathmatchstarts(i)
            P_SpawnPlayer spot
            Exit Sub
        End If
    Next

    ' no good spot, so the player will probably get stuck
    spot = playerstarts(playernum)
    P_SpawnPlayer spot
End Sub

' G_DoReborn
Private Sub G_DoReborn(ByVal playernum As Long)
    Dim i As Long, spot As mapthing_t

    If Not netgame Then
        ' reload the level from scratch
        gameaction = ga_loadlevel
    Else
        ' respawn at the start

        ' first dissasociate the corpse
        mobjs(players(playernum).mo).player = -1

        ' spawn at random spot if in death match
        If deathmatch <> 0 Then
            G_DeathMatchSpawnPlayer playernum
            Exit Sub
        End If

        spot = playerstarts(playernum)
        If G_CheckSpot(playernum, spot) Then
            P_SpawnPlayer spot
            Exit Sub
        End If

        ' try to spawn at one of the other players spots
        For i = 0 To MAXPLAYERS - 1
            spot = playerstarts(i)
            If G_CheckSpot(playernum, spot) Then
                playerstarts(i).typ = playernum + 1 ' fake as other player
                spot = playerstarts(i)
                P_SpawnPlayer spot
                playerstarts(i).typ = i + 1      ' restore
                Exit Sub
            End If
            ' he's going to be inside something.  Too bad.
        Next
        spot = playerstarts(playernum)
        P_SpawnPlayer spot
    End If
End Sub

Public Sub G_ScreenShot()
    gameaction = ga_screenshot
End Sub

Private Sub InitPars()
    Dim v As Variant, i As Long
    If parsInit Then Exit Sub
    ' DOOM Par Times
    v = Array(0, 30, 75, 120, 90, 165, 180, 180, 30, 165)
    For i = 0 To 9: pars(1, i) = v(i): Next
    v = Array(0, 90, 90, 90, 120, 90, 360, 240, 30, 170)
    For i = 0 To 9: pars(2, i) = v(i): Next
    v = Array(0, 90, 45, 90, 150, 90, 90, 165, 30, 135)
    For i = 0 To 9: pars(3, i) = v(i): Next
    ' DOOM II Par Times
    v = Array(30, 90, 120, 120, 90, 150, 120, 120, 270, 90, _
              210, 150, 150, 150, 210, 150, 420, 150, 210, 150, _
              240, 150, 180, 150, 150, 300, 330, 420, 300, 180, _
              120, 30)
    For i = 0 To 31: cpars(i) = v(i): Next
    parsInit = True
End Sub

' G_DoCompleted
Public Sub G_ExitLevel()
    secretexit = False
    gameaction = ga_completed
End Sub

' Here's for the german edition.
Public Sub G_SecretExitLevel()
    ' IF NO WOLF3D LEVELS, NO SECRET EXIT!
    If gamemode = commercial And W_CheckNumForName("map31") < 0 Then
        secretexit = False
    Else
        secretexit = True
    End If
    gameaction = ga_completed
End Sub

Private Sub G_DoCompleted()
    Dim i As Long, j As Long

    InitPars
    gameaction = ga_nothing

    For i = 0 To MAXPLAYERS - 1
        If playeringame(i) Then G_PlayerFinishLevel i ' take away cards and stuff
    Next

    If automapactive Then AM_Stop

    If gamemode <> commercial Then
        Select Case gamemap
            Case 8
                gameaction = ga_victory
                Exit Sub
            Case 9
                For i = 0 To MAXPLAYERS - 1
                    players(i).didsecret = True
                Next
        End Select
    End If

    wminfo.didsecret = players(consoleplayer).didsecret
    wminfo.epsd = gameepisode - 1
    wminfo.last = gamemap - 1

    ' wminfo.next is 0 biased, unlike gamemap
    If gamemode = commercial Then
        If secretexit Then
            Select Case gamemap
                Case 15: wminfo.nxt = 30
                Case 31: wminfo.nxt = 31
            End Select
        Else
            Select Case gamemap
                Case 31, 32: wminfo.nxt = 15
                Case Else: wminfo.nxt = gamemap
            End Select
        End If
    Else
        If secretexit Then
            wminfo.nxt = 8                       ' go to secret level
        ElseIf gamemap = 9 Then
            ' returning from secret level
            Select Case gameepisode
                Case 1: wminfo.nxt = 3
                Case 2: wminfo.nxt = 5
                Case 3: wminfo.nxt = 6
                Case 4: wminfo.nxt = 2
            End Select
        Else
            wminfo.nxt = gamemap                 ' go to next level
        End If
    End If

    wminfo.maxkills = totalkills
    wminfo.maxitems = totalitems
    wminfo.maxsecret = totalsecret
    wminfo.maxfrags = 0
    If gamemode = commercial Then
        wminfo.partime = 35 * cpars(gamemap - 1)
    ElseIf gameepisode <= 3 Then
        wminfo.partime = 35 * pars(gameepisode, gamemap)
    Else
        wminfo.partime = 0                       ' (C reads past the table for episode 4)
    End If
    wminfo.pnum = consoleplayer

    For i = 0 To MAXPLAYERS - 1
        wminfo.plyr(i).inGame = playeringame(i)
        wminfo.plyr(i).skills = players(i).killcount
        wminfo.plyr(i).sitems = players(i).itemcount
        wminfo.plyr(i).ssecret = players(i).secretcount
        wminfo.plyr(i).stime = leveltime
        For j = 0 To 3
            wminfo.plyr(i).frags(j) = players(i).frags(j)
        Next
    Next

    gamestate = GS_INTERMISSION
    viewactive = False
    automapactive = False

    WI_Start wminfo
End Sub

' G_WorldDone
Public Sub G_WorldDone()
    gameaction = ga_worlddone

    If secretexit Then players(consoleplayer).didsecret = True

    If gamemode = commercial Then
        Select Case gamemap
            Case 15, 31
                If secretexit Then F_StartFinale
            Case 6, 11, 20, 30
                F_StartFinale
        End Select
    End If
End Sub

Private Sub G_DoWorldDone()
    gamestate = GS_LEVEL
    gamemap = wminfo.nxt + 1
    G_DoLoadLevel
    gameaction = ga_nothing
    viewactive = True
End Sub

' G_InitFromSavegame - can be called by the startup code or the menu task.
Public Sub G_LoadGame(ByVal name As String)
    savename = name
    gameaction = ga_loadgame
End Sub

Public Sub G_DoLoadGame()
    Dim f As Integer, hdr() As Byte, vcheck As String, i As Long
    Dim a As Long, b As Long, c As Long, marker As Byte, sk As Long, ep As Long, mp As Long

    gameaction = ga_nothing

    If Dir$(savename) = "" Then Exit Sub
    f = FreeFile
    Open savename For Binary Access Read As #f

    ReDim hdr(0 To SAVESTRINGSIZE + VERSIONSIZE + 3 + MAXPLAYERS + 3 - 1)
    Get #f, 1, hdr

    vcheck = ""
    For i = 0 To VERSIONSIZE - 1
        If hdr(SAVESTRINGSIZE + i) = 0 Then Exit For
        vcheck = vcheck & Chr$(hdr(SAVESTRINGSIZE + i))
    Next
    If vcheck <> "version " & VERSION Then
        Close #f
        Exit Sub                                 ' bad version
    End If

    i = SAVESTRINGSIZE + VERSIONSIZE
    sk = hdr(i): ep = hdr(i + 1): mp = hdr(i + 2)
    i = i + 3
    For a = 0 To MAXPLAYERS - 1
        playeringame(a) = (hdr(i + a) <> 0)
    Next
    i = i + MAXPLAYERS

    ' load a base level
    G_InitNew sk, ep, mp

    ' get the times
    a = hdr(i): b = hdr(i + 1): c = hdr(i + 2)
    leveltime = (a * 65536) + (b * 256) + c

    ' dearchive all the modifications
    P_UnArchivePlayers f
    P_UnArchiveWorld f
    P_UnArchiveThinkers f
    P_UnArchiveSpecials f

    Get #f, , marker
    Close #f
    If marker <> &H1D Then I_Error "Bad savegame"

    If setsizeneeded Then R_ExecuteSetViewSize

    ' draw the pattern into the back screen
    R_FillBackScreen
End Sub

' G_SaveGame - called by the menu task.  Description is a 24 byte text string
Public Sub G_SaveGame(ByVal slot As Long, ByVal description As String)
    savegameslot = slot
    savedescription = description
    sendsave = True
End Sub

Private Sub G_DoSaveGame()
    Dim nm As String, f As Integer, hdr() As Byte, i As Long, s As String, marker As Byte

    nm = DoomExePath & SAVEGAMENAME & savegameslot & ".dsg"

    ReDim hdr(0 To SAVESTRINGSIZE + VERSIONSIZE + 3 + MAXPLAYERS + 3 - 1)
    For i = 1 To Len(savedescription)
        If i > SAVESTRINGSIZE Then Exit For
        hdr(i - 1) = Asc(Mid$(savedescription, i, 1))
    Next
    s = "version " & VERSION
    For i = 1 To Len(s)
        hdr(SAVESTRINGSIZE + i - 1) = Asc(Mid$(s, i, 1))
    Next
    i = SAVESTRINGSIZE + VERSIONSIZE
    hdr(i) = gameskill: hdr(i + 1) = gameepisode: hdr(i + 2) = gamemap
    i = i + 3
    If playeringame(0) Then hdr(i) = 1
    If playeringame(1) Then hdr(i + 1) = 1
    If playeringame(2) Then hdr(i + 2) = 1
    If playeringame(3) Then hdr(i + 3) = 1
    i = i + MAXPLAYERS
    hdr(i) = (leveltime \ 65536) And 255
    hdr(i + 1) = (leveltime \ 256) And 255
    hdr(i + 2) = leveltime And 255

    On Error Resume Next
    If Dir$(nm) <> "" Then Kill nm
    On Error GoTo 0
    f = FreeFile
    Open nm For Binary Access Write As #f
    Put #f, 1, hdr

    P_ArchivePlayers f
    P_ArchiveWorld f
    P_ArchiveThinkers f
    P_ArchiveSpecials f

    marker = &H1D                                ' consistancy marker
    Put #f, , marker
    Close #f

    gameaction = ga_nothing
    savedescription = ""

    players(consoleplayer).message = GGSAVED

    ' draw the pattern into the back screen
    R_FillBackScreen
End Sub

' G_InitNew - can be called by the startup code or the menu task,
' consoleplayer, displayplayer, playeringame[] should be set.
Public Sub G_DeferedInitNew(ByVal skill As Long, ByVal episode As Long, ByVal map As Long)
    d_skill = skill
    d_episode = episode
    d_map = map
    gameaction = ga_newgame
End Sub

Private Sub G_DoNewGame()
    demoplayback = False
    netdemo = False
    netgame = False
    deathmatch = 0
    playeringame(1) = False
    playeringame(2) = False
    playeringame(3) = False
    respawnparm = False
    fastparm = False
    nomonsters = False
    consoleplayer = 0
    G_InitNew d_skill, d_episode, d_map
    gameaction = ga_nothing
End Sub

Public Sub G_InitNew(ByVal skill As Long, ByVal episode As Long, ByVal map As Long)
    Dim i As Long

    If paused Then
        paused = False
        S_ResumeSound
    End If

    If skill > sk_nightmare Then skill = sk_nightmare

    ' This was quite messy with SPECIAL and commented parts.
    ' Supposedly hacks to make the latest edition work.
    ' It might not work properly.
    If episode < 1 Then episode = 1

    If gamemode = retail Then
        If episode > 4 Then episode = 4
    ElseIf gamemode = shareware Then
        If episode > 1 Then episode = 1          ' only start episode 1 on shareware
    Else
        If episode > 3 Then episode = 3
    End If

    If map < 1 Then map = 1

    If map > 9 And gamemode <> commercial Then map = 9

    M_ClearRandom

    If skill = sk_nightmare Or respawnparm Then
        respawnmonsters = True
    Else
        respawnmonsters = False
    End If

    If fastparm Or (skill = sk_nightmare And gameskill <> sk_nightmare) Then
        For i = S_SARG_RUN1 To S_SARG_PAIN2
            states(i).tics = SAR(states(i).tics, 1)
        Next
        mobjinfo(MT_BRUISERSHOT).speed = 20 * FRACUNIT
        mobjinfo(MT_HEADSHOT).speed = 20 * FRACUNIT
        mobjinfo(MT_TROOPSHOT).speed = 20 * FRACUNIT
    ElseIf skill <> sk_nightmare And gameskill = sk_nightmare Then
        For i = S_SARG_RUN1 To S_SARG_PAIN2
            states(i).tics = states(i).tics * 2
        Next
        mobjinfo(MT_BRUISERSHOT).speed = 15 * FRACUNIT
        mobjinfo(MT_HEADSHOT).speed = 10 * FRACUNIT
        mobjinfo(MT_TROOPSHOT).speed = 10 * FRACUNIT
    End If

    ' force players to be initialized upon first level load
    For i = 0 To MAXPLAYERS - 1
        players(i).playerstate = PST_REBORN
    Next

    usergame = True                              ' will be set false if a demo
    paused = False
    demoplayback = False
    automapactive = False
    viewactive = True
    gameepisode = episode
    gamemap = map
    gameskill = skill

    viewactive = True

    ' set the sky map for the episode
    If gamemode = commercial Then
        skytexture = R_TextureNumForName("SKY3")
        If gamemap < 12 Then
            skytexture = R_TextureNumForName("SKY1")
        ElseIf gamemap < 21 Then
            skytexture = R_TextureNumForName("SKY2")
        End If
    Else
        Select Case episode
            Case 1: skytexture = R_TextureNumForName("SKY1")
            Case 2: skytexture = R_TextureNumForName("SKY2")
            Case 3: skytexture = R_TextureNumForName("SKY3")
            Case 4: skytexture = R_TextureNumForName("SKY4")
        End Select
    End If

    G_DoLoadLevel
End Sub

' DEMO RECORDING

Private Sub G_ReadDemoTiccmd(cmd As ticcmd_t)
    Dim v As Long
    If demodata(demo_p) = DEMOMARKER Then
        ' end of demo data stream
        G_CheckDemoStatus
        Exit Sub
    End If
    v = demodata(demo_p): If v > 127 Then v = v - 256
    cmd.forwardmove = v
    v = demodata(demo_p + 1): If v > 127 Then v = v - 256
    cmd.sidemove = v
    cmd.angleturn = ToShort(CLng(demodata(demo_p + 2)) * 256)
    cmd.buttons = demodata(demo_p + 3)
    demo_p = demo_p + 4
End Sub

Private Sub G_WriteDemoTiccmd(cmd As ticcmd_t)
    If gamekeydown(Asc("q")) Then G_CheckDemoStatus ' press q to end demo recording
    If Not demorecording Then Exit Sub
    demodata(demo_p) = cmd.forwardmove And 255
    demodata(demo_p + 1) = cmd.sidemove And 255
    demodata(demo_p + 2) = SAR(cmd.angleturn + 128, 8) And 255
    demodata(demo_p + 3) = cmd.buttons And 255

    If demo_p > demoend - 16 Then
        ' no more space
        G_CheckDemoStatus
        Exit Sub
    End If

    G_ReadDemoTiccmd cmd                         ' make SURE it is exactly the same
End Sub

' G_RecordDemo
Public Sub G_RecordDemo(ByVal name As String)
    Dim i As Long, maxsize As Long

    usergame = False
    demoname = name & ".lmp"
    maxsize = &H20000
    i = M_CheckParm("-maxdemo")
    If i > 0 And i < myargc - 1 Then maxsize = Val(myargv(i + 1)) * 1024
    ReDim demodata(0 To maxsize - 1)
    demoend = maxsize
    demorecording = True
End Sub

Public Sub G_BeginRecording()
    Dim i As Long

    demo_p = 0

    demodata(demo_p) = VERSION: demo_p = demo_p + 1
    demodata(demo_p) = gameskill: demo_p = demo_p + 1
    demodata(demo_p) = gameepisode: demo_p = demo_p + 1
    demodata(demo_p) = gamemap: demo_p = demo_p + 1
    demodata(demo_p) = deathmatch: demo_p = demo_p + 1
    demodata(demo_p) = IIf(respawnparm, 1, 0): demo_p = demo_p + 1
    demodata(demo_p) = IIf(fastparm, 1, 0): demo_p = demo_p + 1
    demodata(demo_p) = IIf(nomonsters, 1, 0): demo_p = demo_p + 1
    demodata(demo_p) = consoleplayer: demo_p = demo_p + 1

    For i = 0 To MAXPLAYERS - 1
        demodata(demo_p) = IIf(playeringame(i), 1, 0)
        demo_p = demo_p + 1
    Next
End Sub

' G_PlayDemo
Public Sub G_DeferedPlayDemo(ByVal name As String)
    defdemoname = name
    gameaction = ga_playdemo
End Sub

Public Sub G_PlayDemo(ByVal name As String)
    G_DeferedPlayDemo name
End Sub

Private Sub G_DoPlayDemo()
    Dim skill As Long, i As Long, episode As Long, map As Long, lump As Long, ver As Long

    gameaction = ga_nothing
    lump = W_CheckNumForName(defdemoname)
    If lump < 0 Then
        I_Printf "Demo " & defdemoname & " not found"
        Exit Sub
    End If
    W_ReadLump lump, demodata
    demo_p = 0
    ver = demodata(demo_p): demo_p = demo_p + 1
    ' (DOOM 1.9 demos are version 109; the game logic is the same)
    If ver <> VERSION And ver <> 109 Then
        I_Printf "Demo is from a different game version! (" & ver & ")"
        gameaction = ga_nothing
        Exit Sub
    End If

    skill = demodata(demo_p): demo_p = demo_p + 1
    episode = demodata(demo_p): demo_p = demo_p + 1
    map = demodata(demo_p): demo_p = demo_p + 1
    deathmatch = demodata(demo_p): demo_p = demo_p + 1
    respawnparm = (demodata(demo_p) <> 0): demo_p = demo_p + 1
    fastparm = (demodata(demo_p) <> 0): demo_p = demo_p + 1
    nomonsters = (demodata(demo_p) <> 0): demo_p = demo_p + 1
    consoleplayer = demodata(demo_p): demo_p = demo_p + 1

    For i = 0 To MAXPLAYERS - 1
        playeringame(i) = (demodata(demo_p) <> 0)
        demo_p = demo_p + 1
    Next
    If playeringame(1) Then
        netgame = True
        netdemo = True
    End If

    ' don't spend a lot of time in loadlevel
    precache = False
    G_InitNew skill, episode, map
    precache = True

    usergame = False
    demoplayback = True
End Sub

' G_TimeDemo
Public Sub G_TimeDemo(ByVal name As String)
    nodrawers = (M_CheckParm("-nodraw") > 0)
    noblit = (M_CheckParm("-noblit") > 0)
    timingdemo = True
    singletics = True

    defdemoname = name
    gameaction = ga_playdemo
End Sub

' G_CheckDemoStatus - called after a death or level completion to allow
' demos to be cleaned up.  Returns true if a new demo loop action will take
' place.
Public Function G_CheckDemoStatus() As Boolean
    Dim endtime As Long, f As Integer, buf() As Byte

    If timingdemo Then
        endtime = I_GetTime()
        I_Printf "timed " & gametic & " gametics in " & (endtime - starttime) & " realtics (" & _
                 Format$(gametic * 35# / IIf(endtime - starttime > 0, endtime - starttime, 1), "0.0") & " fps)"
        If M_CheckParm("-hidden") = 0 Then
            MsgBox "timed " & gametic & " gametics in " & (endtime - starttime) & " realtics", vbInformation, "DoomVB"
        End If
        I_Quit
    End If

    If demoplayback Then
        If singledemo Then I_Quit

        demoplayback = False
        netdemo = False
        netgame = False
        deathmatch = 0
        playeringame(1) = False
        playeringame(2) = False
        playeringame(3) = False
        respawnparm = False
        fastparm = False
        nomonsters = False
        consoleplayer = 0
        D_AdvanceDemo
        G_CheckDemoStatus = True
        Exit Function
    End If

    If demorecording Then
        demodata(demo_p) = DEMOMARKER
        demo_p = demo_p + 1
        ReDim buf(0 To demo_p - 1)
        CopyMemory buf(0), demodata(0), demo_p
        M_WriteFile demoname, buf, demo_p
        demorecording = False
        I_Error "Demo " & demoname & " recorded"
    End If

    G_CheckDemoStatus = False
End Function
