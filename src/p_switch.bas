Attribute VB_Name = "p_switch"
' p_switch.bas - switches, buttons.  Two-state animation.  Exits (p_switch.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private switchlist(0 To 2 * MAXSWITCHES) As Long
Private numswitches As Long

' P_InitSwitchList - only called at game initialization.
Public Sub P_InitSwitchList()
    Dim i As Long, index As Long, episode As Long, lst As Variant

    ' CHANGE THE TEXTURE OF A WALL SWITCH TO ITS OPPOSITE
    lst = Array( _
        "SW1BRCOM", "SW2BRCOM", 1, "SW1BRN1", "SW2BRN1", 1, "SW1BRN2", "SW2BRN2", 1, _
        "SW1BRNGN", "SW2BRNGN", 1, "SW1BROWN", "SW2BROWN", 1, "SW1COMM", "SW2COMM", 1, _
        "SW1COMP", "SW2COMP", 1, "SW1DIRT", "SW2DIRT", 1, "SW1EXIT", "SW2EXIT", 1, _
        "SW1GRAY", "SW2GRAY", 1, "SW1GRAY1", "SW2GRAY1", 1, "SW1METAL", "SW2METAL", 1, _
        "SW1PIPE", "SW2PIPE", 1, "SW1SLAD", "SW2SLAD", 1, "SW1STARG", "SW2STARG", 1, _
        "SW1STON1", "SW2STON1", 1, "SW1STON2", "SW2STON2", 1, "SW1STONE", "SW2STONE", 1, _
        "SW1STRTN", "SW2STRTN", 1, _
        "SW1BLUE", "SW2BLUE", 2, "SW1CMT", "SW2CMT", 2, "SW1GARG", "SW2GARG", 2, _
        "SW1GSTON", "SW2GSTON", 2, "SW1HOT", "SW2HOT", 2, "SW1LION", "SW2LION", 2, _
        "SW1SATYR", "SW2SATYR", 2, "SW1SKIN", "SW2SKIN", 2, "SW1VINE", "SW2VINE", 2, _
        "SW1WOOD", "SW2WOOD", 2, _
        "SW1PANEL", "SW2PANEL", 3, "SW1ROCK", "SW2ROCK", 3, "SW1MET2", "SW2MET2", 3, _
        "SW1WDMET", "SW2WDMET", 3, "SW1BRIK", "SW2BRIK", 3, "SW1MOD1", "SW2MOD1", 3, _
        "SW1ZIM", "SW2ZIM", 3, "SW1STON6", "SW2STON6", 3, "SW1TEK", "SW2TEK", 3, _
        "SW1MARB", "SW2MARB", 3, "SW1SKULL", "SW2SKULL", 3)

    episode = 1

    If gamemode = registered Or gamemode = retail Then
        episode = 2
    ElseIf gamemode = commercial Then
        episode = 3
    End If

    index = 0
    For i = 0 To (UBound(lst) + 1) \ 3 - 1
        If lst(i * 3 + 2) <= episode Then
            switchlist(index) = R_TextureNumForName(lst(i * 3))
            index = index + 1
            switchlist(index) = R_TextureNumForName(lst(i * 3 + 1))
            index = index + 1
        End If
    Next
    numswitches = index \ 2
    switchlist(index) = -1
End Sub

' Start a button counting down till it turns off.
Private Sub P_StartButton(ByVal line As Long, ByVal w As Long, ByVal texture As Long, ByVal time As Long)
    Dim i As Long

    ' See if button is already pressed
    For i = 0 To MAXBUTTONS - 1
        If buttonlist(i).btimer <> 0 And buttonlist(i).linenum = line Then Exit Sub
    Next

    For i = 0 To MAXBUTTONS - 1
        If buttonlist(i).btimer = 0 Then
            buttonlist(i).linenum = line
            buttonlist(i).where = w
            buttonlist(i).btexture = texture
            buttonlist(i).btimer = time
            buttonlist(i).soundorg = SectorSoundOrg(lines(line).frontsector)
            Exit Sub
        End If
    Next

    I_Error "P_StartButton: no button slots left!"
End Sub

' Function that changes wall texture.  Tell it if switch is ok to use again
' (1=yes, it's a button).
Public Sub P_ChangeSwitchTexture(ByVal line As Long, ByVal useAgain As Long)
    Dim texTop As Long, texMid As Long, texBot As Long, i As Long, sound As Long, sd As Long

    If useAgain = 0 Then lines(line).special = 0

    sd = lines(line).sidenum(0)
    texTop = sides(sd).toptexture
    texMid = sides(sd).midtexture
    texBot = sides(sd).bottomtexture

    sound = sfx_swtchn

    ' EXIT SWITCH?
    If lines(line).special = 11 Then sound = sfx_swtchx

    For i = 0 To numswitches * 2 - 1
        If switchlist(i) = texTop Then
            S_StartSound buttonlist(0).soundorg, sound    ' (sic: C uses buttonlist->soundorg)
            sides(sd).toptexture = switchlist(i Xor 1)

            If useAgain <> 0 Then P_StartButton line, bw_top, switchlist(i), BUTTONTIME

            Exit Sub
        ElseIf switchlist(i) = texMid Then
            S_StartSound buttonlist(0).soundorg, sound
            sides(sd).midtexture = switchlist(i Xor 1)

            If useAgain <> 0 Then P_StartButton line, bw_middle, switchlist(i), BUTTONTIME

            Exit Sub
        ElseIf switchlist(i) = texBot Then
            S_StartSound buttonlist(0).soundorg, sound
            sides(sd).bottomtexture = switchlist(i Xor 1)

            If useAgain <> 0 Then P_StartButton line, bw_bottom, switchlist(i), BUTTONTIME

            Exit Sub
        End If
    Next
End Sub

' P_UseSpecialLine - called when a thing uses a special line.  Only the
' front sides of lines are usable.
Public Function P_UseSpecialLine(ByVal thing As Long, ByVal line As Long, ByVal side As Long) As Boolean
    ' Err...
    ' Use the back sides of VERY SPECIAL lines...
    If side <> 0 Then
        Select Case lines(line).special
            Case 124
                ' Sliding door open&close
                ' UNUSED?
            Case Else
                P_UseSpecialLine = False
                Exit Function
        End Select
    End If

    ' Switches that other things can activate.
    If mobjs(thing).player < 0 Then
        ' never open secret doors
        If (lines(line).flags And ML_SECRET) <> 0 Then
            P_UseSpecialLine = False
            Exit Function
        End If

        Select Case lines(line).special
            Case 1, 32, 33, 34
                ' MANUAL DOOR RAISE, MANUAL BLUE, MANUAL RED, MANUAL YELLOW
            Case Else
                P_UseSpecialLine = False
                Exit Function
        End Select
    End If

    ' do something
    Select Case lines(line).special
        ' MANUALS
        Case 1, 26, 27, 28, 31, 32, 33, 34, 117, 118
            ' Vertical Door / Blue Door/Locked / Yellow / Red / Manual opens /
            ' Blue Card / Red Card / Yellow Card / Blazing door raise / open
            EV_VerticalDoor line, thing

        ' SWITCHES
        Case 7
            ' Build Stairs
            If EV_BuildStairs(line, build8) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 9
            ' Change Donut
            If EV_DoDonut(line) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 11
            ' Exit level
            P_ChangeSwitchTexture line, 0
            G_ExitLevel
        Case 14
            ' Raise Floor 32 and change texture
            If EV_DoPlat(line, raiseAndChange, 32) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 15
            ' Raise Floor 24 and change texture
            If EV_DoPlat(line, raiseAndChange, 24) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 18
            ' Raise Floor to next highest floor
            If EV_DoFloor(line, raiseFloorToNearest) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 20
            ' Raise Plat next highest floor and change texture
            If EV_DoPlat(line, raiseToNearestAndChange, 0) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 21
            ' PlatDownWaitUpStay
            If EV_DoPlat(line, downWaitUpStay, 0) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 23
            ' Lower Floor to Lowest
            If EV_DoFloor(line, lowerFloorToLowest) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 29
            ' Raise Door
            If EV_DoDoor(line, vld_normal) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 41
            ' Lower Ceiling to Floor
            If EV_DoCeiling(line, lowerToFloor) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 71
            ' Turbo Lower Floor
            If EV_DoFloor(line, turboLower) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 49
            ' Ceiling Crush And Raise
            If EV_DoCeiling(line, crushAndRaise) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 50
            ' Close Door
            If EV_DoDoor(line, vld_close) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 51
            ' Secret EXIT
            P_ChangeSwitchTexture line, 0
            G_SecretExitLevel
        Case 55
            ' Raise Floor Crush
            If EV_DoFloor(line, raiseFloorCrush) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 101
            ' Raise Floor
            If EV_DoFloor(line, raiseFloor) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 102
            ' Lower Floor to Surrounding floor height
            If EV_DoFloor(line, lowerFloor) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 103
            ' Open Door
            If EV_DoDoor(line, vld_open) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 111
            ' Blazing Door Raise (faster than TURBO!)
            If EV_DoDoor(line, vld_blazeRaise) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 112
            ' Blazing Door Open (faster than TURBO!)
            If EV_DoDoor(line, vld_blazeOpen) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 113
            ' Blazing Door Close (faster than TURBO!)
            If EV_DoDoor(line, vld_blazeClose) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 122
            ' Blazing PlatDownWaitUpStay
            If EV_DoPlat(line, blazeDWUS, 0) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 127
            ' Build Stairs Turbo 16
            If EV_BuildStairs(line, turbo16) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 131
            ' Raise Floor Turbo
            If EV_DoFloor(line, raiseFloorTurbo) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 133, 135, 137
            ' BlzOpenDoor BLUE / RED / YELLOW
            If EV_DoLockedDoor(line, vld_blazeOpen, thing) <> 0 Then P_ChangeSwitchTexture line, 0
        Case 140
            ' Raise Floor 512
            If EV_DoFloor(line, raiseFloor512) <> 0 Then P_ChangeSwitchTexture line, 0

        ' BUTTONS
        Case 42
            ' Close Door
            If EV_DoDoor(line, vld_close) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 43
            ' Lower Ceiling to Floor
            If EV_DoCeiling(line, lowerToFloor) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 45
            ' Lower Floor to Surrounding floor height
            If EV_DoFloor(line, lowerFloor) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 60
            ' Lower Floor to Lowest
            If EV_DoFloor(line, lowerFloorToLowest) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 61
            ' Open Door
            If EV_DoDoor(line, vld_open) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 62
            ' PlatDownWaitUpStay
            If EV_DoPlat(line, downWaitUpStay, 1) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 63
            ' Raise Door
            If EV_DoDoor(line, vld_normal) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 64
            ' Raise Floor to ceiling
            If EV_DoFloor(line, raiseFloor) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 66
            ' Raise Floor 24 and change texture
            If EV_DoPlat(line, raiseAndChange, 24) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 67
            ' Raise Floor 32 and change texture
            If EV_DoPlat(line, raiseAndChange, 32) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 65
            ' Raise Floor Crush
            If EV_DoFloor(line, raiseFloorCrush) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 68
            ' Raise Plat to next highest floor and change texture
            If EV_DoPlat(line, raiseToNearestAndChange, 0) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 69
            ' Raise Floor to next highest floor
            If EV_DoFloor(line, raiseFloorToNearest) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 70
            ' Turbo Lower Floor
            If EV_DoFloor(line, turboLower) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 114
            ' Blazing Door Raise (faster than TURBO!)
            If EV_DoDoor(line, vld_blazeRaise) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 115
            ' Blazing Door Open (faster than TURBO!)
            If EV_DoDoor(line, vld_blazeOpen) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 116
            ' Blazing Door Close (faster than TURBO!)
            If EV_DoDoor(line, vld_blazeClose) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 123
            ' Blazing PlatDownWaitUpStay
            If EV_DoPlat(line, blazeDWUS, 0) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 132
            ' Raise Floor Turbo
            If EV_DoFloor(line, raiseFloorTurbo) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 99, 134, 136
            ' BlzOpenDoor BLUE / RED / YELLOW
            If EV_DoLockedDoor(line, vld_blazeOpen, thing) <> 0 Then P_ChangeSwitchTexture line, 1
        Case 138
            ' Light Turn On
            EV_LightTurnOn line, 255
            P_ChangeSwitchTexture line, 1
        Case 139
            ' Light Turn Off
            EV_LightTurnOn line, 35
            P_ChangeSwitchTexture line, 1
    End Select

    P_UseSpecialLine = True
End Function
