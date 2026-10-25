Attribute VB_Name = "d_main"
' d_main.bas - DOOM main program: startup, main loop, display, title loop
' (d_main.c; the network game parts are in d_net and i_net).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
'
' Extra command line options of this port (besides the original ones):
'   -iwad <file>      main WAD to use (default: search the exe folder)
'   -scale <n>        window size multiplier (default 2 = 640x480)
'   -width <pixels>   window client width
'   -nomouse          do not capture the mouse
'   -hidden           run without showing the window (automated tests)
'   -shot <tic> <bmp> save a screenshot at the given gametic
'   -quitat <tic>     quit at the given gametic
'   -exitat <tic>     exit the level at the given gametic (as the exit switch)
'   -key <tic>[+<n>] <key>  press a key at the given gametic and release it
'                     one (or n) tics later; key is a letter, digit or esc
'                     enter space bksp tab up down left right minus equals
'                     shift ctrl alt pause f1..f12
'   -netdrop <n>      drop every n-th network packet sent (tests resending)
'   -net hosts may be written host:port (several players on one computer)
Option Explicit

Public demosequence As Long
Public pagetic As Long
Public pagename As String
Public singletics As Boolean

Private shotTics() As Long
Private shotFiles() As String
Private numShots As Long
Private quitAtTic As Long
Private exitAtTic As Long
Private keyTics() As Long
Private keyUpTics() As Long
Private keyCodes() As Long
Private keyPhase() As Long                       ' 0 = press pending, 1 = release pending, 2 = done
Private numKeys As Long

' D_PostEvent - called by the I/O functions when input is detected
Public Sub D_PostEvent(ev As event_t)
    events(eventhead) = ev
    eventhead = (eventhead + 1) And (MAXEVENTS - 1)
End Sub

' D_ProcessEvents - send all the events of the given timestamp down the responder chain
Public Sub D_ProcessEvents()
    If gamemode = commercial Then
        If W_CheckNumForName("map01") < 0 Then Exit Sub
    End If
    Do While eventtail <> eventhead
        If Not M_Responder(events(eventtail)) Then   ' menu ate the event?
            G_Responder events(eventtail)
        End If
        eventtail = (eventtail + 1) And (MAXEVENTS - 1)
    Loop
End Sub

' D_Display - draw current display, possibly wiping it from the previous
Public Sub D_Display()
    Static viewactivestate As Boolean
    Static menuactivestate As Boolean
    Static inhelpscreensstate As Boolean
    Static fullscreen As Boolean
    Static oldgamestate As Long
    Static borderdrawcount As Long
    Static initialised As Boolean
    Dim nowtime As Long, tics As Long, wipestart As Long, y As Long
    Dim done As Boolean, wipe As Boolean, redrawsbar As Boolean

    If Not initialised Then
        oldgamestate = -1
        initialised = True
    End If

    If nodrawers Then Exit Sub                   ' for comparative timing / profiling

    redrawsbar = False

    ' change the view size if needed
    If setsizeneeded Then
        R_ExecuteSetViewSize
        oldgamestate = -1                        ' force background redraw
        borderdrawcount = 3
    End If

    ' save the current screen if about to wipe
    If gamestate <> wipegamestate Then
        wipe = True
        wipe_StartScreen 0, 0, SCREENWIDTH, SCREENHEIGHT
    Else
        wipe = False
    End If

    If gamestate = GS_LEVEL And gametic <> 0 Then HU_Erase

    ' do buffered drawing
    dbgPhase = "D_Display"
    Select Case gamestate
        Case GS_LEVEL
            If gametic <> 0 Then
                If automapactive Then AM_Drawer
                If wipe Or (viewheight <> 200 And fullscreen) Then redrawsbar = True
                If inhelpscreensstate And Not inhelpscreens Then redrawsbar = True ' just put away the help screen
                ST_Drawer (viewheight = 200), redrawsbar
                fullscreen = (viewheight = 200)
            End If
        Case GS_INTERMISSION
            WI_Drawer
        Case GS_FINALE
            F_Drawer
        Case GS_DEMOSCREEN
            D_PageDrawer
    End Select

    ' draw buffered stuff to screen
    I_UpdateNoBlit

    ' draw the view directly
    If gamestate = GS_LEVEL And Not automapactive And gametic <> 0 Then
        dbgPhase = "R_RenderPlayerView"
        R_RenderPlayerView displayplayer
    End If

    If gamestate = GS_LEVEL And gametic <> 0 Then HU_Drawer

    ' clean up border stuff
    If gamestate <> oldgamestate And gamestate <> GS_LEVEL Then I_SetPalette W_CacheLumpName("PLAYPAL")

    ' see if the border needs to be initially drawn
    If gamestate = GS_LEVEL And oldgamestate <> GS_LEVEL Then
        viewactivestate = False                  ' view was not active
        R_FillBackScreen                         ' draw the pattern into the back screen
    End If

    ' see if the border needs to be updated to the screen
    If gamestate = GS_LEVEL And Not automapactive And scaledviewwidth <> 320 Then
        If menuactive Or menuactivestate Or Not viewactivestate Then borderdrawcount = 3
        If borderdrawcount <> 0 Then
            R_DrawViewBorder                     ' erase old menu stuff
            borderdrawcount = borderdrawcount - 1
        End If
    End If

    menuactivestate = menuactive
    viewactivestate = viewactive
    inhelpscreensstate = inhelpscreens
    oldgamestate = gamestate
    wipegamestate = gamestate

    ' draw pause pic
    If paused Then
        If automapactive Then y = 4 Else y = viewwindowy + 4
        V_DrawPatchDirect viewwindowx + (scaledviewwidth - 68) \ 2, y, 0, W_GetNumForName("M_PAUSE")
    End If

    ' menus go directly to the screen
    dbgPhase = "M_Drawer"
    M_Drawer                                     ' menu is drawn even on top of everything

    CheckScreenshotHook

    ' normal update
    If Not wipe Then
        I_FinishUpdate                           ' page flip or blit buffer
        Exit Sub
    End If

    ' wipe update
    wipe_EndScreen 0, 0, SCREENWIDTH, SCREENHEIGHT

    wipestart = I_GetTime() - 1
    Do
        Do
            nowtime = I_GetTime()
            tics = nowtime - wipestart
            If tics = 0 Then
                Sleep 1
                I_UpdateSound                    ' (Win32: no sound interrupt)
            End If
        Loop While tics = 0
        wipestart = nowtime
        done = wipe_ScreenWipe(wipe_Melt, 0, 0, SCREENWIDTH, SCREENHEIGHT, tics)
        I_UpdateNoBlit
        M_Drawer                                 ' menu is drawn even on top of wipes
        I_FinishUpdate                           ' page flip or blit buffer
        I_UpdateSound                            ' (Win32: no sound interrupt)
        I_UpdateMusic
        DoEvents
        If quitRequested Then I_Quit
    Loop While Not done
End Sub

' test hooks: -shot <tic> <file> and -quitat <tic>
Private Sub ParseTestHooks()
    Dim i As Long
    numShots = 0
    For i = 1 To myargc - 3
        If StrComp(myargv(i), "-shot", vbTextCompare) = 0 Then
            ReDim Preserve shotTics(0 To numShots)
            ReDim Preserve shotFiles(0 To numShots)
            shotTics(numShots) = Val(myargv(i + 1))
            shotFiles(numShots) = myargv(i + 2)
            numShots = numShots + 1
        End If
    Next
    i = M_CheckParm("-quitat")
    If i > 0 And i < myargc - 1 Then quitAtTic = Val(myargv(i + 1)) Else quitAtTic = -1
    i = M_CheckParm("-exitat")
    If i > 0 And i < myargc - 1 Then exitAtTic = Val(myargv(i + 1)) Else exitAtTic = -1
    numKeys = 0
    For i = 1 To myargc - 3
        If StrComp(myargv(i), "-key", vbTextCompare) = 0 Then
            ReDim Preserve keyTics(0 To numKeys)
            ReDim Preserve keyUpTics(0 To numKeys)
            ReDim Preserve keyCodes(0 To numKeys)
            ReDim Preserve keyPhase(0 To numKeys)
            keyTics(numKeys) = Val(myargv(i + 1))
            If InStr(myargv(i + 1), "+") > 0 Then
                keyUpTics(numKeys) = keyTics(numKeys) + Val(Mid$(myargv(i + 1), InStr(myargv(i + 1), "+") + 1))
            Else
                keyUpTics(numKeys) = keyTics(numKeys) + 1
            End If
            keyCodes(numKeys) = KeyCodeForName(myargv(i + 2))
            keyPhase(numKeys) = 0
            numKeys = numKeys + 1
        End If
    Next
End Sub

' key code for a -key name
Private Function KeyCodeForName(ByVal nm As String) As Long
    nm = LCase$(nm)
    Select Case nm
        Case "esc": KeyCodeForName = KEY_ESCAPE
        Case "enter": KeyCodeForName = KEY_ENTER
        Case "space": KeyCodeForName = 32
        Case "bksp": KeyCodeForName = KEY_BACKSPACE
        Case "tab": KeyCodeForName = KEY_TAB
        Case "up": KeyCodeForName = KEY_UPARROW
        Case "down": KeyCodeForName = KEY_DOWNARROW
        Case "left": KeyCodeForName = KEY_LEFTARROW
        Case "right": KeyCodeForName = KEY_RIGHTARROW
        Case "minus": KeyCodeForName = KEY_MINUS
        Case "equals": KeyCodeForName = KEY_EQUALS
        Case "shift": KeyCodeForName = KEY_RSHIFT
        Case "ctrl": KeyCodeForName = KEY_RCTRL
        Case "alt": KeyCodeForName = KEY_RALT
        Case "pause": KeyCodeForName = KEY_PAUSE
        Case "f11": KeyCodeForName = KEY_F11
        Case "f12": KeyCodeForName = KEY_F12
        Case "f1", "f2", "f3", "f4", "f5", "f6", "f7", "f8", "f9", "f10"
            KeyCodeForName = KEY_F1 + Val(Mid$(nm, 2)) - 1
        Case Else
            KeyCodeForName = Asc(nm)
    End Select
End Function

' posts the -key presses and releases that are due
Private Sub CheckKeyHook()
    Dim i As Long, ev As event_t
    For i = 0 To numKeys - 1
        If keyPhase(i) = 0 And gametic >= keyTics(i) Then
            ev.typ = ev_keydown
            ev.data1 = keyCodes(i)
            D_PostEvent ev
            keyPhase(i) = 1
        ElseIf keyPhase(i) = 1 And gametic >= keyUpTics(i) Then
            ev.typ = ev_keyup
            ev.data1 = keyCodes(i)
            D_PostEvent ev
            keyPhase(i) = 2
        End If
    Next
End Sub

Private Sub CheckScreenshotHook()
    Dim i As Long
    For i = 0 To numShots - 1
        If shotTics(i) >= 0 And gametic >= shotTics(i) Then
            M_WriteBMP shotFiles(i)
            shotTics(i) = -1
        End If
    Next
End Sub

' D_DoomLoop
Public Sub D_DoomLoop()
    On Error GoTo crashed

    If demorecording Then G_BeginRecording

    If M_CheckParm("-debugfile") > 0 Then
        debugfile = FreeFile
        Open DoomExePath & "debug" & consoleplayer & ".txt" For Output As #debugfile
        I_Printf "debug output to: debug" & consoleplayer & ".txt"
    End If

    I_InitGraphics

    Do
        ' frame syncronous IO operations
        I_StartFrame

        ' process one or more tics
        If singletics Then
            I_StartTic
            D_ProcessEvents
            G_BuildTiccmd netcmds(consoleplayer, maketic Mod BACKUPTICS)
            If advancedemo Then D_DoAdvanceDemo
            M_Ticker
            G_Ticker
            gametic = gametic + 1
            maketic = maketic + 1
        Else
            TryRunTics                           ' will run at least one tic
        End If

        dbgPhase = "S_UpdateSounds"
        S_UpdateSounds players(consoleplayer).mo ' move positional sounds

        ' Update display, next frame, with current state.
        D_Display

        dbgPhase = "I_UpdateSound"
        I_UpdateSound
        I_UpdateMusic

        If quitAtTic >= 0 And gametic >= quitAtTic Then I_Quit
        If exitAtTic >= 0 And gametic >= exitAtTic And gamestate = GS_LEVEL Then
            exitAtTic = -1
            G_ExitLevel
        End If
        If numKeys > 0 Then CheckKeyHook
    Loop
    Exit Sub

crashed:
    I_RuntimeError "D_DoomLoop (" & dbgPhase & ")"
End Sub

' ------------------------------------------------------------------ title loop
Public Sub D_PageTicker()
    pagetic = pagetic - 1
    If pagetic < 0 Then D_AdvanceDemo
End Sub

Public Sub D_PageDrawer()
    V_DrawPatch 0, 0, 0, W_GetNumForName(pagename)
End Sub

' Called after each demo or intro demosequence finishes
Public Sub D_AdvanceDemo()
    advancedemo = True
End Sub

' This cycles through the demo sequences.
Public Sub D_DoAdvanceDemo()
    players(consoleplayer).playerstate = PST_LIVE ' not reborn
    advancedemo = False
    usergame = False                             ' no save / end game here
    paused = False
    gameaction = ga_nothing

    If gamemode = retail Then
        demosequence = (demosequence + 1) Mod 7
    Else
        demosequence = (demosequence + 1) Mod 6
    End If

    Select Case demosequence
        Case 0
            If gamemode = commercial Then pagetic = 35 * 11 Else pagetic = 170
            gamestate = GS_DEMOSCREEN
            pagename = "TITLEPIC"
            If gamemode = commercial Then S_StartMusic mus_dm2ttl Else S_StartMusic mus_intro
        Case 1
            G_DeferedPlayDemo "demo1"
        Case 2
            pagetic = 200
            gamestate = GS_DEMOSCREEN
            pagename = "CREDIT"
        Case 3
            G_DeferedPlayDemo "demo2"
        Case 4
            gamestate = GS_DEMOSCREEN
            If gamemode = commercial Then
                pagetic = 35 * 11
                pagename = "TITLEPIC"
                S_StartMusic mus_dm2ttl
            Else
                pagetic = 200
                If gamemode = retail Then pagename = "CREDIT" Else pagename = "HELP2"
            End If
        Case 5
            G_DeferedPlayDemo "demo3"
        Case 6
            G_DeferedPlayDemo "demo4"
    End Select
End Sub

Public Sub D_StartTitle()
    gameaction = ga_nothing
    demosequence = -1
    D_AdvanceDemo
End Sub

Public Sub D_AddFile(ByVal file As String)
    ReDim Preserve wadfilenames(0 To numwadfiles)
    wadfilenames(numwadfiles) = file
    numwadfiles = numwadfiles + 1
End Sub

' Checks availability of IWAD files by name, to determine whether
' registered/commercial features should be executed (notably loading PWADs).
Private Sub IdentifyVersion()
    Dim p As Long, iwad As String, cand As Variant, nm As String

    p = M_CheckParm("-iwad")
    If p > 0 And p < myargc - 1 Then
        iwad = myargv(p + 1)
        If InStr(iwad, "\") = 0 And InStr(iwad, ":") = 0 Then
            If Dir$(iwad) = "" And Dir$(DoomExePath & iwad) <> "" Then iwad = DoomExePath & iwad
        End If
        If Dir$(iwad) = "" Then I_Error "IWAD not found: " & iwad
    Else
        For Each cand In Array("doom2f.wad", "doom2.wad", "plutonia.wad", "tnt.wad", "doomu.wad", _
                               "doom.wad", "doom1.wad", "freedoom2.wad", "freedoom1.wad", "freedm.wad")
            If Dir$(DoomExePath & cand) <> "" Then
                iwad = DoomExePath & cand
                Exit For
            End If
        Next
        If iwad = "" Then
            I_Error "Game mode indeterminate." & vbCrLf & vbCrLf & _
                "Put a DOOM IWAD (doom1.wad, doom.wad, doom2.wad, freedoom1.wad, ...)" & vbCrLf & _
                "in " & DoomExePath & " or use -iwad <file>."
        End If
    End If

    nm = LCase$(Mid$(iwad, InStrRev(iwad, "\") + 1))
    gamemission = doom
    Select Case nm
        Case "doom2f.wad"
            gamemode = commercial: gamemission = doom2: language = french
        Case "doom2.wad", "freedoom2.wad", "freedm.wad"
            gamemode = commercial: gamemission = doom2
        Case "plutonia.wad"
            gamemode = commercial: gamemission = pack_plut
        Case "tnt.wad"
            gamemode = commercial: gamemission = pack_tnt
        Case "doomu.wad"
            gamemode = retail
        Case "doom.wad"
            gamemode = registered
        Case "doom1.wad"
            gamemode = shareware
        Case Else
            gamemode = indetermined                ' decided from the lumps below
    End Select
    D_AddFile iwad
End Sub

' For IWADs not recognised by name (and doom.wad, which is also the Ultimate
' DOOM file name), decide the game mode from the map lumps present.
Private Sub RefineVersionFromLumps()
    If gamemode = indetermined Or gamemode = registered Then
        If W_CheckNumForName("MAP01") >= 0 Then
            gamemode = commercial
            If gamemission = doom Then gamemission = doom2
        ElseIf W_CheckNumForName("E4M1") >= 0 Then
            gamemode = retail
        ElseIf W_CheckNumForName("E2M1") >= 0 Then
            gamemode = registered
        Else
            gamemode = shareware
        End If
    End If
End Sub

Private Sub InitStatics()
    InitTables
    InitInfoTables
    InitSoundTables
    InitStrings
    InitWeaponInfo
    maxammo(0) = 200: maxammo(1) = 50: maxammo(2) = 300: maxammo(3) = 50
    clipammo(0) = 10: clipammo(1) = 4: clipammo(2) = 20: clipammo(3) = 1
    player_names(0) = HUSTR_PLRGREEN
    player_names(1) = HUSTR_PLRINDIGO
    player_names(2) = HUSTR_PLRBROWN
    player_names(3) = HUSTR_PLRRED
    wipegamestate = GS_DEMOSCREEN
    viewactive = True
    ReDim wadfilenames(0 To 0)
    numwadfiles = 0
End Sub

Public Sub D_DoomMain()
    Dim p As Long, file As String, sc As Long

    On Error GoTo crashed

    InitStatics
    IdentifyVersion
    ParseTestHooks

    modifiedgame = False

    nomonsters = (M_CheckParm("-nomonsters") > 0)
    respawnparm = (M_CheckParm("-respawn") > 0)
    fastparm = (M_CheckParm("-fast") > 0)
    devparm = (M_CheckParm("-devparm") > 0)
    checklinks = (M_CheckParm("-checklinks") > 0)
    p = M_CheckParm("-watch")
    If p > 0 And p < myargc - 1 Then watchslot = Val(myargv(p + 1)) Else watchslot = -1
    p = M_CheckParm("-synclog")
    If p > 0 And p < myargc - 1 Then
        synclog = FreeFile
        Open myargv(p + 1) For Output As #synclog
    End If
    If M_CheckParm("-altdeath") > 0 Then
        deathmatch = 2
    ElseIf M_CheckParm("-deathmatch") > 0 Then
        deathmatch = 1
    End If

    I_Printf "DoomVB - Visual Basic 6 port of DOOM v" & (VERSION \ 100) & "." & (VERSION Mod 100)
    If devparm Then I_Printf D_DEVSTR

    basedefault = DoomExePath & "default.cfg"
    p = M_CheckParm("-config")
    If p > 0 And p < myargc - 1 Then basedefault = myargv(p + 1)

    ' turbo option
    p = M_CheckParm("-turbo")
    If p > 0 Then
        sc = 200
        If p < myargc - 1 Then sc = Val(myargv(p + 1))
        If sc < 10 Then sc = 10
        If sc > 400 Then sc = 400
        I_Printf "turbo scale: " & sc & "%"
        G_SetTurbo sc
    End If

    p = M_CheckParm("-file")
    If p > 0 Then
        ' the parms after p are wadfile/lump names, until end of parms or another - preceded parm
        modifiedgame = True                      ' homebrew levels
        p = p + 1
        Do While p < myargc
            If Left$(myargv(p), 1) = "-" Then Exit Do
            D_AddFile myargv(p)
            p = p + 1
        Loop
    End If

    p = M_CheckParm("-playdemo")
    If p = 0 Then p = M_CheckParm("-timedemo")
    If p > 0 And p < myargc - 1 Then
        file = myargv(p + 1)
        If LCase$(Right$(file, 4)) <> ".lmp" Then file = file & ".lmp"
        If Dir$(file) <> "" Then
            D_AddFile file
            I_Printf "Playing demo " & file
        End If
    End If

    ' get skill / episode / map from parms
    startskill = sk_medium
    startepisode = 1
    startmap = 1
    autostart = False

    p = M_CheckParm("-skill")
    If p > 0 And p < myargc - 1 Then
        startskill = Asc(myargv(p + 1)) - Asc("1")
        autostart = True
    End If

    p = M_CheckParm("-episode")
    If p > 0 And p < myargc - 1 Then
        startepisode = Asc(myargv(p + 1)) - Asc("0")
        startmap = 1
        autostart = True
    End If

    p = M_CheckParm("-warp")
    If p > 0 And p < myargc - 1 Then
        If gamemode = commercial Then
            startmap = Val(myargv(p + 1))
        ElseIf p < myargc - 2 Then
            startepisode = Asc(myargv(p + 1)) - Asc("0")
            startmap = Asc(myargv(p + 2)) - Asc("0")
        End If
        autostart = True
    End If

    ' init subsystems
    I_Printf "V_Init: allocate screens."
    V_Init

    I_Printf "M_LoadDefaults: Load system defaults."
    M_LoadDefaults                               ' load before initing other systems

    I_Printf "W_Init: Init WADfiles."
    W_InitMultipleFiles wadfilenames, numwadfiles
    RefineVersionFromLumps

    Select Case gamemode
        Case retail: I_Printf "The Ultimate DOOM Startup"
        Case shareware: I_Printf "DOOM Shareware Startup"
        Case registered: I_Printf "DOOM Registered Startup"
        Case commercial: I_Printf "DOOM 2: Hell on Earth Startup"
        Case Else: I_Printf "Public DOOM"
    End Select

    I_Printf "M_Init: Init miscellaneous info."
    dbgPhase = "M_Init"
    M_Init

    I_Printf "R_Init: Init DOOM refresh daemon"
    dbgPhase = "R_Init"
    R_Init

    I_Printf "P_Init: Init Playloop state."
    dbgPhase = "P_Init"
    P_Init

    I_Printf "I_Init: Setting up machine state."
    dbgPhase = "I_Init"
    I_Init

    I_Printf "D_CheckNetGame: Checking network game status."
    D_CheckNetGame

    I_Printf "S_Init: Setting up sound."
    dbgPhase = "S_Init"
    S_Init snd_SfxVolume, snd_MusicVolume

    I_Printf "HU_Init: Setting up heads up display."
    dbgPhase = "HU_Init"
    HU_Init

    I_Printf "ST_Init: Init status bar."
    dbgPhase = "ST_Init"
    ST_Init
    dbgPhase = ""

    ' start the apropriate game based on parms
    p = M_CheckParm("-record")
    If p > 0 And p < myargc - 1 Then
        G_RecordDemo myargv(p + 1)
        autostart = True
    End If

    p = M_CheckParm("-playdemo")
    If p > 0 And p < myargc - 1 Then
        singledemo = True                        ' quit after one demo
        G_DeferedPlayDemo myargv(p + 1)
        D_DoomLoop                               ' never returns
    End If

    p = M_CheckParm("-timedemo")
    If p > 0 And p < myargc - 1 Then
        G_TimeDemo myargv(p + 1)
        D_DoomLoop                               ' never returns
    End If

    p = M_CheckParm("-loadgame")
    If p > 0 And p < myargc - 1 Then
        file = DoomExePath & SAVEGAMENAME & myargv(p + 1) & ".dsg"
        G_LoadGame file
    End If

    If gameaction <> ga_loadgame Then
        If autostart Or netgame Then
            G_InitNew startskill, startepisode, startmap
        Else
            D_StartTitle                         ' start up intro loop
        End If
    End If

    D_DoomLoop                                   ' never returns
    Exit Sub

crashed:
    I_RuntimeError "D_DoomMain (" & dbgPhase & ")"
End Sub
