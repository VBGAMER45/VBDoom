Attribute VB_Name = "p_user"
' p_user.bas - player related stuff.  Bobbing POV/weapon, movement.
' Pending weapon (p_user.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' Index of the special effects (INVUL inverse) map.
Private Const INVERSECOLORMAP As Long = 32

' Movement.

' 16 pixels of bob
Private Const MAXBOB As Long = &H100000

Private Const ANG5 As Long = ANG90 \ 18

Private onground As Boolean

' P_Thrust - moves the given origin along a given angle.
Private Sub P_Thrust(ByVal player As Long, ByVal angle As Long, ByVal move As Long)
    Dim mo As Long
    mo = players(player).mo
    angle = FINEANGLE(angle)

    mobjs(mo).momx = mobjs(mo).momx + FixedMul(move, finecosine(angle))
    mobjs(mo).momy = mobjs(mo).momy + FixedMul(move, finesine(angle))
End Sub

' P_CalcHeight - calculate the walking / running height adjustment
Private Sub P_CalcHeight(ByVal player As Long)
    Dim angle As Long, bob As Long, mo As Long

    mo = players(player).mo

    ' Regular movement bobbing (needs to be calculated for gun swing even if
    ' not on ground)
    ' OPTIMIZE: tablify angle
    ' Note: a LUT allows for effects like a ramp with low health.
    players(player).bob = FixedMul(mobjs(mo).momx, mobjs(mo).momx) + FixedMul(mobjs(mo).momy, mobjs(mo).momy)

    players(player).bob = SAR(players(player).bob, 2)

    If players(player).bob > MAXBOB Then players(player).bob = MAXBOB

    If (players(player).cheats And CF_NOMOMENTUM) <> 0 Or Not onground Then
        players(player).viewz = mobjs(mo).z + PLAYER_VIEWHEIGHT

        If players(player).viewz > mobjs(mo).ceilingz - 4 * FRACUNIT Then
            players(player).viewz = mobjs(mo).ceilingz - 4 * FRACUNIT
        End If

        players(player).viewz = mobjs(mo).z + players(player).viewheight
        Exit Sub
    End If

    angle = ((FINEANGLES \ 20) * leveltime) And FINEMASK
    bob = FixedMul(players(player).bob \ 2, finesine(angle))

    ' move viewheight
    If players(player).playerstate = PST_LIVE Then
        players(player).viewheight = players(player).viewheight + players(player).deltaviewheight

        If players(player).viewheight > PLAYER_VIEWHEIGHT Then
            players(player).viewheight = PLAYER_VIEWHEIGHT
            players(player).deltaviewheight = 0
        End If

        If players(player).viewheight < PLAYER_VIEWHEIGHT \ 2 Then
            players(player).viewheight = PLAYER_VIEWHEIGHT \ 2
            If players(player).deltaviewheight <= 0 Then players(player).deltaviewheight = 1
        End If

        If players(player).deltaviewheight <> 0 Then
            players(player).deltaviewheight = players(player).deltaviewheight + FRACUNIT \ 4
            If players(player).deltaviewheight = 0 Then players(player).deltaviewheight = 1
        End If
    End If
    players(player).viewz = mobjs(mo).z + players(player).viewheight + bob

    If players(player).viewz > mobjs(mo).ceilingz - 4 * FRACUNIT Then
        players(player).viewz = mobjs(mo).ceilingz - 4 * FRACUNIT
    End If
End Sub

' P_MovePlayer
Private Sub P_MovePlayer(ByVal player As Long)
    Dim mo As Long

    mo = players(player).mo

    mobjs(mo).angle = mobjs(mo).angle + (players(player).cmd.angleturn * &H10000)

    ' Do not let the player control movement if not onground.
    onground = (mobjs(mo).z <= mobjs(mo).floorz)

    If players(player).cmd.forwardmove <> 0 And onground Then
        P_Thrust player, mobjs(mo).angle, players(player).cmd.forwardmove * 2048
    End If

    If players(player).cmd.sidemove <> 0 And onground Then
        P_Thrust player, mobjs(mo).angle - ANG90, players(player).cmd.sidemove * 2048
    End If

    If (players(player).cmd.forwardmove <> 0 Or players(player).cmd.sidemove <> 0) And _
       mobjs(mo).state = S_PLAY Then
        P_SetMobjState mo, S_PLAY_RUN1
    End If
End Sub

' P_DeathThink - fall on your face when dying.  Decrease POV height to floor
' height.
Private Sub P_DeathThink(ByVal player As Long)
    Dim angle As Long, delta As Long, mo As Long, att As Long

    P_MovePsprites player

    ' fall to the ground
    If players(player).viewheight > 6 * FRACUNIT Then
        players(player).viewheight = players(player).viewheight - FRACUNIT
    End If

    If players(player).viewheight < 6 * FRACUNIT Then players(player).viewheight = 6 * FRACUNIT

    players(player).deltaviewheight = 0
    mo = players(player).mo
    onground = (mobjs(mo).z <= mobjs(mo).floorz)
    P_CalcHeight player

    att = players(player).attacker
    If att <> 0 And att <> mo Then
        angle = R_PointToAngle2(mobjs(mo).x, mobjs(mo).y, mobjs(att).x, mobjs(att).y)

        delta = angle - mobjs(mo).angle

        If ULT(delta, ANG5) Or UGT(delta, -ANG5) Then
            ' Looking at killer, so fade damage flash down.
            mobjs(mo).angle = angle

            If players(player).damagecount <> 0 Then players(player).damagecount = players(player).damagecount - 1
        ElseIf ULT(delta, ANG180) Then
            mobjs(mo).angle = mobjs(mo).angle + ANG5
        Else
            mobjs(mo).angle = mobjs(mo).angle - ANG5
        End If
    ElseIf players(player).damagecount <> 0 Then
        players(player).damagecount = players(player).damagecount - 1
    End If

    If (players(player).cmd.buttons And BT_USE) <> 0 Then players(player).playerstate = PST_REBORN
End Sub

' P_PlayerThink
Public Sub P_PlayerThink(ByVal player As Long)
    Dim newweapon As Long, mo As Long

    mo = players(player).mo

    ' fixme: do this in the cheat code
    If (players(player).cheats And CF_NOCLIP) <> 0 Then
        mobjs(mo).flags = mobjs(mo).flags Or MF_NOCLIP
    Else
        mobjs(mo).flags = mobjs(mo).flags And (Not MF_NOCLIP)
    End If

    ' chain saw run forward
    If (mobjs(mo).flags And MF_JUSTATTACKED) <> 0 Then
        players(player).cmd.angleturn = 0
        players(player).cmd.forwardmove = &HC800& \ 512
        players(player).cmd.sidemove = 0
        mobjs(mo).flags = mobjs(mo).flags And (Not MF_JUSTATTACKED)
    End If

    If players(player).playerstate = PST_DEAD Then
        P_DeathThink player
        Exit Sub
    End If

    ' Move around.
    ' Reactiontime is used to prevent movement for a bit after a teleport.
    If mobjs(mo).reactiontime <> 0 Then
        mobjs(mo).reactiontime = mobjs(mo).reactiontime - 1
    Else
        P_MovePlayer player
    End If

    P_CalcHeight player

    If sectors(subsectors(mobjs(mo).subsector).sector).special <> 0 Then P_PlayerInSpecialSector player

    ' Check for weapon change.

    ' A special event has no other buttons.
    If (players(player).cmd.buttons And BT_SPECIAL) <> 0 Then players(player).cmd.buttons = 0

    If (players(player).cmd.buttons And BT_CHANGE) <> 0 Then
        ' The actual changing of the weapon is done when the weapon psprite
        ' can do it (read: not in the middle of an attack).
        newweapon = (players(player).cmd.buttons And BT_WEAPONMASK) \ 8

        If newweapon = wp_fist And players(player).weaponowned(wp_chainsaw) And _
           Not (players(player).readyweapon = wp_chainsaw And players(player).powers(pw_strength) <> 0) Then
            newweapon = wp_chainsaw
        End If

        If gamemode = commercial And newweapon = wp_shotgun And _
           players(player).weaponowned(wp_supershotgun) And players(player).readyweapon <> wp_supershotgun Then
            newweapon = wp_supershotgun
        End If

        If players(player).weaponowned(newweapon) And newweapon <> players(player).readyweapon Then
            ' Do not go to plasma or BFG in shareware, even if cheated.
            If (newweapon <> wp_plasma And newweapon <> wp_bfg) Or gamemode <> shareware Then
                players(player).pendingweapon = newweapon
            End If
        End If
    End If

    ' check for use
    If (players(player).cmd.buttons And BT_USE) <> 0 Then
        If Not players(player).usedown Then
            P_UseLines player
            players(player).usedown = True
        End If
    Else
        players(player).usedown = False
    End If

    ' cycle psprites
    P_MovePsprites player

    ' Counters, time dependend power ups.

    ' Strength counts up to diminish fade.
    If players(player).powers(pw_strength) <> 0 Then
        players(player).powers(pw_strength) = players(player).powers(pw_strength) + 1
    End If

    If players(player).powers(pw_invulnerability) <> 0 Then
        players(player).powers(pw_invulnerability) = players(player).powers(pw_invulnerability) - 1
    End If

    If players(player).powers(pw_invisibility) <> 0 Then
        players(player).powers(pw_invisibility) = players(player).powers(pw_invisibility) - 1
        If players(player).powers(pw_invisibility) = 0 Then
            mobjs(mo).flags = mobjs(mo).flags And (Not MF_SHADOW)
        End If
    End If

    If players(player).powers(pw_infrared) <> 0 Then
        players(player).powers(pw_infrared) = players(player).powers(pw_infrared) - 1
    End If

    If players(player).powers(pw_ironfeet) <> 0 Then
        players(player).powers(pw_ironfeet) = players(player).powers(pw_ironfeet) - 1
    End If

    If players(player).damagecount <> 0 Then players(player).damagecount = players(player).damagecount - 1

    If players(player).bonuscount <> 0 Then players(player).bonuscount = players(player).bonuscount - 1

    ' Handling colormaps.
    If players(player).powers(pw_invulnerability) <> 0 Then
        If players(player).powers(pw_invulnerability) > 4 * 32 Or (players(player).powers(pw_invulnerability) And 8) <> 0 Then
            players(player).fixedcolormap = INVERSECOLORMAP
        Else
            players(player).fixedcolormap = 0
        End If
    ElseIf players(player).powers(pw_infrared) <> 0 Then
        If players(player).powers(pw_infrared) > 4 * 32 Or (players(player).powers(pw_infrared) And 8) <> 0 Then
            ' almost full bright
            players(player).fixedcolormap = 1
        Else
            players(player).fixedcolormap = 0
        End If
    Else
        players(player).fixedcolormap = 0
    End If
End Sub
