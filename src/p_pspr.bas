Attribute VB_Name = "p_pspr"
' p_pspr.bas - weapon sprite animation, weapon objects.  Action functions for
' weapons (p_pspr.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const LOWERSPEED As Long = FRACUNIT * 6
Private Const RAISESPEED As Long = FRACUNIT * 6

Private Const WEAPONBOTTOM As Long = 128 * FRACUNIT
Private Const WEAPONTOP As Long = 32 * FRACUNIT

' plasma cells for a bfg attack
Private Const BFGCELLS As Long = 40

Private swingx As Long
Private swingy As Long
Private bulletslope As Long

Private Function P_SubRandom() As Long
    Dim r As Long
    r = P_Random()
    P_SubRandom = r - P_Random()
End Function

' Weapon action dispatcher (C: state->action.acp2 (player, psp)).
Private Sub P_CallPspAction(ByVal action As Long, ByVal player As Long, ByVal psp As Long)
    Select Case action
        Case AF_WeaponReady: A_WeaponReady player, psp
        Case AF_ReFire: A_ReFire player, psp
        Case AF_CheckReload: A_CheckReload player, psp
        Case AF_Lower: A_Lower player, psp
        Case AF_Raise: A_Raise player, psp
        Case AF_GunFlash: A_GunFlash player, psp
        Case AF_Punch: A_Punch player, psp
        Case AF_Saw: A_Saw player, psp
        Case AF_FireMissile: A_FireMissile player, psp
        Case AF_FireBFG: A_FireBFG player, psp
        Case AF_FirePlasma: A_FirePlasma player, psp
        Case AF_FirePistol: A_FirePistol player, psp
        Case AF_FireShotgun: A_FireShotgun player, psp
        Case AF_FireShotgun2: A_FireShotgun2 player, psp
        Case AF_FireCGun: A_FireCGun player, psp
        Case AF_Light0: players(player).extralight = 0
        Case AF_Light1: players(player).extralight = 1
        Case AF_Light2: players(player).extralight = 2
        Case AF_BFGsound: S_StartSound players(player).mo, sfx_bfg
        Case AF_OpenShotgun2: S_StartSound players(player).mo, sfx_dbopn
        Case AF_LoadShotgun2: S_StartSound players(player).mo, sfx_dbload
        Case AF_CloseShotgun2
            S_StartSound players(player).mo, sfx_dbcls
            A_ReFire player, psp
        Case Else
            I_Error "P_CallPspAction: unexpected action " & action
    End Select
End Sub

' P_SetPsprite
Public Sub P_SetPsprite(ByVal player As Long, ByVal position As Long, ByVal stnum As Long)
    Do
        If stnum = 0 Then
            ' object removed itself
            players(player).psprites(position).state = S_NULL
            Exit Do
        End If

        players(player).psprites(position).state = stnum
        players(player).psprites(position).tics = states(stnum).tics ' could be 0

        If states(stnum).misc1 <> 0 Then
            ' coordinate set
            players(player).psprites(position).sx = states(stnum).misc1 * FRACUNIT
            players(player).psprites(position).sy = states(stnum).misc2 * FRACUNIT
        End If

        ' Call action routine.
        ' Modified handling.
        If states(stnum).action <> AF_NULL Then
            P_CallPspAction states(stnum).action, player, position
            If players(player).psprites(position).state = S_NULL Then Exit Do
        End If

        stnum = states(players(player).psprites(position).state).nextstate
    Loop While players(player).psprites(position).tics = 0
    ' an initial state of 0 could cycle through
End Sub

' P_CalcSwing
Private Sub P_CalcSwing(ByVal player As Long)
    Dim swing As Long, angle As Long

    ' OPTIMIZE: tablify this.
    ' A LUT would allow for different modes, and add flexibility.
    swing = players(player).bob

    angle = ((FINEANGLES \ 70) * leveltime) And FINEMASK
    swingx = FixedMul(swing, finesine(angle))

    angle = ((FINEANGLES \ 70) * leveltime + FINEANGLES \ 2) And FINEMASK
    swingy = -FixedMul(swingx, finesine(angle))
End Sub

' P_BringUpWeapon - starts bringing the pending weapon up from the bottom of
' the screen.  Uses player
Private Sub P_BringUpWeapon(ByVal player As Long)
    Dim newstate As Long

    If players(player).pendingweapon = wp_nochange Then players(player).pendingweapon = players(player).readyweapon

    If players(player).pendingweapon = wp_chainsaw Then S_StartSound players(player).mo, sfx_sawup

    newstate = weaponinfo(players(player).pendingweapon).upstate

    players(player).pendingweapon = wp_nochange
    players(player).psprites(ps_weapon).sy = WEAPONBOTTOM

    P_SetPsprite player, ps_weapon, newstate
End Sub

' P_CheckAmmo - returns true if there is enough ammo to shoot.  If not,
' selects the next weapon to use.
Private Function P_CheckAmmo(ByVal player As Long) As Boolean
    Dim ammo As Long, count As Long

    ammo = weaponinfo(players(player).readyweapon).ammo

    ' Minimal amount for one shot varies.
    If players(player).readyweapon = wp_bfg Then
        count = BFGCELLS
    ElseIf players(player).readyweapon = wp_supershotgun Then
        count = 2                                ' Double barrel.
    Else
        count = 1                                ' Regular.
    End If

    ' Some do not need ammunition anyway.
    ' Return if current ammunition sufficient.
    If ammo = am_noammo Then
        P_CheckAmmo = True
        Exit Function
    End If
    If players(player).ammo(ammo) >= count Then
        P_CheckAmmo = True
        Exit Function
    End If

    ' Out of ammo, pick a weapon to change to.
    ' Preferences are set here.
    Do
        If players(player).weaponowned(wp_plasma) And players(player).ammo(am_cell) <> 0 And gamemode <> shareware Then
            players(player).pendingweapon = wp_plasma
        ElseIf players(player).weaponowned(wp_supershotgun) And players(player).ammo(am_shell) > 2 And gamemode = commercial Then
            players(player).pendingweapon = wp_supershotgun
        ElseIf players(player).weaponowned(wp_chaingun) And players(player).ammo(am_clip) <> 0 Then
            players(player).pendingweapon = wp_chaingun
        ElseIf players(player).weaponowned(wp_shotgun) And players(player).ammo(am_shell) <> 0 Then
            players(player).pendingweapon = wp_shotgun
        ElseIf players(player).ammo(am_clip) <> 0 Then
            players(player).pendingweapon = wp_pistol
        ElseIf players(player).weaponowned(wp_chainsaw) Then
            players(player).pendingweapon = wp_chainsaw
        ElseIf players(player).weaponowned(wp_missile) And players(player).ammo(am_misl) <> 0 Then
            players(player).pendingweapon = wp_missile
        ElseIf players(player).weaponowned(wp_bfg) And players(player).ammo(am_cell) > 40 And gamemode <> shareware Then
            players(player).pendingweapon = wp_bfg
        Else
            ' If everything fails.
            players(player).pendingweapon = wp_fist
        End If
    Loop While players(player).pendingweapon = wp_nochange

    ' Now set appropriate weapon overlay.
    P_SetPsprite player, ps_weapon, weaponinfo(players(player).readyweapon).downstate

    P_CheckAmmo = False
End Function

' P_FireWeapon.
Private Sub P_FireWeapon(ByVal player As Long)
    Dim newstate As Long

    If Not P_CheckAmmo(player) Then Exit Sub

    P_SetMobjState players(player).mo, S_PLAY_ATK1
    newstate = weaponinfo(players(player).readyweapon).atkstate
    P_SetPsprite player, ps_weapon, newstate
    P_NoiseAlert players(player).mo, players(player).mo
End Sub

' P_DropWeapon - player died, so put the weapon away.
Public Sub P_DropWeapon(ByVal player As Long)
    P_SetPsprite player, ps_weapon, weaponinfo(players(player).readyweapon).downstate
End Sub

' A_WeaponReady - the player can fire the weapon or change to another
' weapon at this time.  Follows after getting weapon up, or after previous
' attack/fire sequence.
Private Sub A_WeaponReady(ByVal player As Long, ByVal psp As Long)
    Dim newstate As Long, angle As Long, mo As Long

    mo = players(player).mo

    ' get out of attack state
    If mobjs(mo).state = S_PLAY_ATK1 Or mobjs(mo).state = S_PLAY_ATK2 Then P_SetMobjState mo, S_PLAY

    If players(player).readyweapon = wp_chainsaw And players(player).psprites(psp).state = S_SAW Then
        S_StartSound mo, sfx_sawidl
    End If

    ' check for change
    '  if player is dead, put the weapon away
    If players(player).pendingweapon <> wp_nochange Or players(player).health = 0 Then
        ' change weapon
        '  (pending weapon should allready be validated)
        newstate = weaponinfo(players(player).readyweapon).downstate
        P_SetPsprite player, ps_weapon, newstate
        Exit Sub
    End If

    ' check for fire
    '  the missile launcher and bfg do not auto fire
    If (players(player).cmd.buttons And BT_ATTACK) <> 0 Then
        If Not players(player).attackdown Or _
           (players(player).readyweapon <> wp_missile And players(player).readyweapon <> wp_bfg) Then
            players(player).attackdown = True
            P_FireWeapon player
            Exit Sub
        End If
    Else
        players(player).attackdown = False
    End If

    ' bob the weapon based on movement speed
    angle = (128 * leveltime) And FINEMASK
    players(player).psprites(psp).sx = FRACUNIT + FixedMul(players(player).bob, finecosine(angle))
    angle = angle And (FINEANGLES \ 2 - 1)
    players(player).psprites(psp).sy = WEAPONTOP + FixedMul(players(player).bob, finesine(angle))
End Sub

' A_ReFire - the player can re-fire the weapon without lowering it entirely.
Private Sub A_ReFire(ByVal player As Long, ByVal psp As Long)
    ' check for fire
    '  (if a weaponchange is pending, let it go through instead)
    If (players(player).cmd.buttons And BT_ATTACK) <> 0 And _
       players(player).pendingweapon = wp_nochange And players(player).health <> 0 Then
        players(player).refire = players(player).refire + 1
        P_FireWeapon player
    Else
        players(player).refire = 0
        P_CheckAmmo player
    End If
End Sub

Private Sub A_CheckReload(ByVal player As Long, ByVal psp As Long)
    P_CheckAmmo player
End Sub

' A_Lower - lowers current weapon, and changes weapon at bottom.
Private Sub A_Lower(ByVal player As Long, ByVal psp As Long)
    players(player).psprites(psp).sy = players(player).psprites(psp).sy + LOWERSPEED

    ' Is already down.
    If players(player).psprites(psp).sy < WEAPONBOTTOM Then Exit Sub

    ' Player is dead.
    If players(player).playerstate = PST_DEAD Then
        players(player).psprites(psp).sy = WEAPONBOTTOM

        ' don't bring weapon back up
        Exit Sub
    End If

    ' The old weapon has been lowered off the screen, so change the weapon
    ' and start raising it
    If players(player).health = 0 Then
        ' Player is dead, so keep the weapon off screen.
        P_SetPsprite player, ps_weapon, S_NULL
        Exit Sub
    End If

    players(player).readyweapon = players(player).pendingweapon

    P_BringUpWeapon player
End Sub

' A_Raise
Private Sub A_Raise(ByVal player As Long, ByVal psp As Long)
    Dim newstate As Long

    players(player).psprites(psp).sy = players(player).psprites(psp).sy - RAISESPEED

    If players(player).psprites(psp).sy > WEAPONTOP Then Exit Sub

    players(player).psprites(psp).sy = WEAPONTOP

    ' The weapon has been raised all the way, so change to the ready state.
    newstate = weaponinfo(players(player).readyweapon).readystate

    P_SetPsprite player, ps_weapon, newstate
End Sub

' A_GunFlash
Private Sub A_GunFlash(ByVal player As Long, ByVal psp As Long)
    P_SetMobjState players(player).mo, S_PLAY_ATK2
    P_SetPsprite player, ps_flash, weaponinfo(players(player).readyweapon).flashstate
End Sub

' WEAPON ATTACKS

' A_Punch
Private Sub A_Punch(ByVal player As Long, ByVal psp As Long)
    Dim angle As Long, damage As Long, slope As Long, mo As Long

    mo = players(player).mo
    damage = ((P_Random() Mod 10) + 1) * 2

    If players(player).powers(pw_strength) <> 0 Then damage = damage * 10

    angle = mobjs(mo).angle
    angle = angle + P_SubRandom() * &H40000
    slope = P_AimLineAttack(mo, angle, MELEERANGE)
    P_LineAttack mo, angle, MELEERANGE, slope, damage

    ' turn to face target
    If linetarget <> 0 Then
        S_StartSound mo, sfx_punch
        mobjs(mo).angle = R_PointToAngle2(mobjs(mo).x, mobjs(mo).y, mobjs(linetarget).x, mobjs(linetarget).y)
    End If
End Sub

' A_Saw
Private Sub A_Saw(ByVal player As Long, ByVal psp As Long)
    Dim angle As Long, damage As Long, slope As Long, mo As Long, d As Long

    mo = players(player).mo
    damage = 2 * ((P_Random() Mod 10) + 1)
    angle = mobjs(mo).angle
    angle = angle + P_SubRandom() * &H40000

    ' use meleerange + 1 se the puff doesn't skip the flash
    slope = P_AimLineAttack(mo, angle, MELEERANGE + 1)
    P_LineAttack mo, angle, MELEERANGE + 1, slope, damage

    If linetarget = 0 Then
        S_StartSound mo, sfx_sawful
        Exit Sub
    End If
    S_StartSound mo, sfx_sawhit

    ' turn to face target
    angle = R_PointToAngle2(mobjs(mo).x, mobjs(mo).y, mobjs(linetarget).x, mobjs(linetarget).y)
    d = angle - mobjs(mo).angle
    If UGT(d, ANG180) Then
        If ULT(d, -(ANG90 \ 20)) Then
            mobjs(mo).angle = angle + ANG90 \ 21
        Else
            mobjs(mo).angle = mobjs(mo).angle - ANG90 \ 20
        End If
    Else
        If UGT(d, ANG90 \ 20) Then
            mobjs(mo).angle = angle - ANG90 \ 21
        Else
            mobjs(mo).angle = mobjs(mo).angle + ANG90 \ 20
        End If
    End If
    mobjs(mo).flags = mobjs(mo).flags Or MF_JUSTATTACKED
End Sub

' A_FireMissile
Private Sub A_FireMissile(ByVal player As Long, ByVal psp As Long)
    Dim a As Long
    a = weaponinfo(players(player).readyweapon).ammo
    players(player).ammo(a) = players(player).ammo(a) - 1
    P_SpawnPlayerMissile players(player).mo, MT_ROCKET
End Sub

' A_FireBFG
Private Sub A_FireBFG(ByVal player As Long, ByVal psp As Long)
    Dim a As Long
    a = weaponinfo(players(player).readyweapon).ammo
    players(player).ammo(a) = players(player).ammo(a) - BFGCELLS
    P_SpawnPlayerMissile players(player).mo, MT_BFG
End Sub

' A_FirePlasma
Private Sub A_FirePlasma(ByVal player As Long, ByVal psp As Long)
    Dim a As Long
    a = weaponinfo(players(player).readyweapon).ammo
    players(player).ammo(a) = players(player).ammo(a) - 1

    P_SetPsprite player, ps_flash, weaponinfo(players(player).readyweapon).flashstate + (P_Random() And 1)

    P_SpawnPlayerMissile players(player).mo, MT_PLASMA
End Sub

' P_BulletSlope - sets a slope so a near miss is at aproximately the height
' of the intended target
Private Sub P_BulletSlope(ByVal mo As Long)
    Dim an As Long

    ' see which target is to be aimed at
    an = mobjs(mo).angle
    bulletslope = P_AimLineAttack(mo, an, 16 * 64 * FRACUNIT)

    If linetarget = 0 Then
        an = an + &H4000000
        bulletslope = P_AimLineAttack(mo, an, 16 * 64 * FRACUNIT)
        If linetarget = 0 Then
            an = an - &H8000000
            bulletslope = P_AimLineAttack(mo, an, 16 * 64 * FRACUNIT)
        End If
    End If
End Sub

' P_GunShot
Private Sub P_GunShot(ByVal mo As Long, ByVal accurate As Boolean)
    Dim angle As Long, damage As Long

    damage = 5 * ((P_Random() Mod 3) + 1)
    angle = mobjs(mo).angle

    If Not accurate Then angle = angle + P_SubRandom() * &H40000

    P_LineAttack mo, angle, MISSILERANGE, bulletslope, damage
End Sub

' A_FirePistol
Private Sub A_FirePistol(ByVal player As Long, ByVal psp As Long)
    Dim a As Long, mo As Long
    mo = players(player).mo
    S_StartSound mo, sfx_pistol

    P_SetMobjState mo, S_PLAY_ATK2
    a = weaponinfo(players(player).readyweapon).ammo
    players(player).ammo(a) = players(player).ammo(a) - 1

    P_SetPsprite player, ps_flash, weaponinfo(players(player).readyweapon).flashstate

    P_BulletSlope mo
    P_GunShot mo, (players(player).refire = 0)
End Sub

' A_FireShotgun
Private Sub A_FireShotgun(ByVal player As Long, ByVal psp As Long)
    Dim i As Long, a As Long, mo As Long

    mo = players(player).mo
    S_StartSound mo, sfx_shotgn
    P_SetMobjState mo, S_PLAY_ATK2

    a = weaponinfo(players(player).readyweapon).ammo
    players(player).ammo(a) = players(player).ammo(a) - 1

    P_SetPsprite player, ps_flash, weaponinfo(players(player).readyweapon).flashstate

    P_BulletSlope mo

    For i = 0 To 6
        P_GunShot mo, False
    Next
End Sub

' A_FireShotgun2
Private Sub A_FireShotgun2(ByVal player As Long, ByVal psp As Long)
    Dim i As Long, angle As Long, damage As Long, a As Long, mo As Long, sl As Long

    mo = players(player).mo
    S_StartSound mo, sfx_dshtgn
    P_SetMobjState mo, S_PLAY_ATK2

    a = weaponinfo(players(player).readyweapon).ammo
    players(player).ammo(a) = players(player).ammo(a) - 2

    P_SetPsprite player, ps_flash, weaponinfo(players(player).readyweapon).flashstate

    P_BulletSlope mo

    For i = 0 To 19
        damage = 5 * ((P_Random() Mod 3) + 1)
        angle = mobjs(mo).angle
        angle = angle + P_SubRandom() * &H80000
        sl = bulletslope + P_SubRandom() * 32
        P_LineAttack mo, angle, MISSILERANGE, sl, damage
    Next
End Sub

' A_FireCGun
Private Sub A_FireCGun(ByVal player As Long, ByVal psp As Long)
    Dim a As Long, mo As Long

    mo = players(player).mo
    S_StartSound mo, sfx_pistol

    a = weaponinfo(players(player).readyweapon).ammo
    If players(player).ammo(a) = 0 Then Exit Sub

    P_SetMobjState mo, S_PLAY_ATK2
    players(player).ammo(a) = players(player).ammo(a) - 1

    P_SetPsprite player, ps_flash, weaponinfo(players(player).readyweapon).flashstate + _
                 players(player).psprites(psp).state - S_CHAIN1

    P_BulletSlope mo

    P_GunShot mo, (players(player).refire = 0)
End Sub

' A_BFGSpray - spawn a BFG explosion on every monster in view (a mobj action).
Public Sub A_BFGSpray(ByVal mo As Long)
    Dim i As Long, j As Long, damage As Long, an As Long, lt As Long

    ' offset angles from its attack angle
    For i = 0 To 39
        an = mobjs(mo).angle - ANG90 \ 2 + (ANG90 \ 40) * i

        ' mo->target is the originator (player) of the missile
        P_AimLineAttack mobjs(mo).target, an, 16 * 64 * FRACUNIT

        lt = linetarget
        If lt <> 0 Then
            P_SpawnMobj mobjs(lt).x, mobjs(lt).y, mobjs(lt).z + SAR(mobjs(lt).height, 2), MT_EXTRABFG

            damage = 0
            For j = 0 To 14
                damage = damage + (P_Random() And 7) + 1
            Next

            P_DamageMobj lt, mobjs(mo).target, mobjs(mo).target, damage
        End If
    Next
End Sub

' P_SetupPsprites - called at start of level for each player.
Public Sub P_SetupPsprites(ByVal player As Long)
    Dim i As Long

    ' remove all psprites
    For i = 0 To NUMPSPRITES - 1
        players(player).psprites(i).state = S_NULL
    Next

    ' spawn the gun
    players(player).pendingweapon = players(player).readyweapon
    P_BringUpWeapon player
End Sub

' P_MovePsprites - called every tic by player thinking routine.
Public Sub P_MovePsprites(ByVal player As Long)
    Dim i As Long

    For i = 0 To NUMPSPRITES - 1
        ' a null state means not active
        If players(player).psprites(i).state <> S_NULL Then
            ' drop tic count and possibly change state
            ' a -1 tic count never changes
            If players(player).psprites(i).tics <> -1 Then
                players(player).psprites(i).tics = players(player).psprites(i).tics - 1
                If players(player).psprites(i).tics = 0 Then
                    P_SetPsprite player, i, states(players(player).psprites(i).state).nextstate
                End If
            End If
        End If
    Next

    players(player).psprites(ps_flash).sx = players(player).psprites(ps_weapon).sx
    players(player).psprites(ps_flash).sy = players(player).psprites(ps_weapon).sy
End Sub
