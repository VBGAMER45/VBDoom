Attribute VB_Name = "p_inter"
' p_inter.bas - handling interactions (i.e., collisions) (p_inter.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const BONUSADD As Long = 6

' a weapon is found with two clip loads, a big item has five clip loads
' (maxammo / clipammo are initialised in d_main: 200,50,300,50 / 10,4,20,1)

' GET STUFF

' P_GiveAmmo - num is the number of clip loads, not the individual count (0=1/2 clip).
' Returns false if the ammo can't be picked up at all
Private Function P_GiveAmmo(ByVal player As Long, ByVal ammo As Long, ByVal num As Long) As Boolean
    Dim oldammo As Long

    If ammo = am_noammo Then
        P_GiveAmmo = False
        Exit Function
    End If

    If ammo < 0 Or ammo > NUMAMMO Then I_Error "P_GiveAmmo: bad type " & ammo

    If players(player).ammo(ammo) = players(player).maxammo(ammo) Then
        P_GiveAmmo = False
        Exit Function
    End If

    If num <> 0 Then
        num = num * clipammo(ammo)
    Else
        num = clipammo(ammo) \ 2
    End If

    If gameskill = sk_baby Or gameskill = sk_nightmare Then
        ' give double ammo in trainer mode, you'll need in nightmare
        num = num * 2
    End If

    oldammo = players(player).ammo(ammo)
    players(player).ammo(ammo) = players(player).ammo(ammo) + num

    If players(player).ammo(ammo) > players(player).maxammo(ammo) Then
        players(player).ammo(ammo) = players(player).maxammo(ammo)
    End If

    ' If non zero ammo, don't change up weapons, player was lower on purpose.
    If oldammo <> 0 Then
        P_GiveAmmo = True
        Exit Function
    End If

    ' We were down to zero, so select a new weapon.
    ' Preferences are not user selectable.
    Select Case ammo
        Case am_clip
            If players(player).readyweapon = wp_fist Then
                If players(player).weaponowned(wp_chaingun) Then
                    players(player).pendingweapon = wp_chaingun
                Else
                    players(player).pendingweapon = wp_pistol
                End If
            End If
        Case am_shell
            If players(player).readyweapon = wp_fist Or players(player).readyweapon = wp_pistol Then
                If players(player).weaponowned(wp_shotgun) Then players(player).pendingweapon = wp_shotgun
            End If
        Case am_cell
            If players(player).readyweapon = wp_fist Or players(player).readyweapon = wp_pistol Then
                If players(player).weaponowned(wp_plasma) Then players(player).pendingweapon = wp_plasma
            End If
        Case am_misl
            If players(player).readyweapon = wp_fist Then
                If players(player).weaponowned(wp_missile) Then players(player).pendingweapon = wp_missile
            End If
    End Select

    P_GiveAmmo = True
End Function

' P_GiveWeapon - the weapon name may have a MF_DROPPED flag ored in.
Private Function P_GiveWeapon(ByVal player As Long, ByVal weapon As Long, ByVal dropped As Boolean) As Boolean
    Dim gaveammo As Boolean, gaveweapon As Boolean

    If netgame And deathmatch <> 2 And Not dropped Then
        ' leave placed weapons forever on net games
        If players(player).weaponowned(weapon) Then
            P_GiveWeapon = False
            Exit Function
        End If

        players(player).bonuscount = players(player).bonuscount + BONUSADD
        players(player).weaponowned(weapon) = True

        If deathmatch <> 0 Then
            P_GiveAmmo player, weaponinfo(weapon).ammo, 5
        Else
            P_GiveAmmo player, weaponinfo(weapon).ammo, 2
        End If
        players(player).pendingweapon = weapon

        If player = consoleplayer Then S_StartSound 0, sfx_wpnup
        P_GiveWeapon = False
        Exit Function
    End If

    If weaponinfo(weapon).ammo <> am_noammo Then
        ' give one clip with a dropped weapon, two clips with a found weapon
        If dropped Then
            gaveammo = P_GiveAmmo(player, weaponinfo(weapon).ammo, 1)
        Else
            gaveammo = P_GiveAmmo(player, weaponinfo(weapon).ammo, 2)
        End If
    Else
        gaveammo = False
    End If

    If players(player).weaponowned(weapon) Then
        gaveweapon = False
    Else
        gaveweapon = True
        players(player).weaponowned(weapon) = True
        players(player).pendingweapon = weapon
    End If

    P_GiveWeapon = (gaveweapon Or gaveammo)
End Function

' P_GiveBody - returns false if the body isn't needed at all
Private Function P_GiveBody(ByVal player As Long, ByVal num As Long) As Boolean
    If players(player).health >= MAXHEALTH Then
        P_GiveBody = False
        Exit Function
    End If

    players(player).health = players(player).health + num
    If players(player).health > MAXHEALTH Then players(player).health = MAXHEALTH
    mobjs(players(player).mo).health = players(player).health

    P_GiveBody = True
End Function

' P_GiveArmor - returns false if the armor is worse than the current armor.
Private Function P_GiveArmor(ByVal player As Long, ByVal armortype As Long) As Boolean
    Dim hits As Long

    hits = armortype * 100
    If players(player).armorpoints >= hits Then
        P_GiveArmor = False                      ' don't pick up
        Exit Function
    End If

    players(player).armortype = armortype
    players(player).armorpoints = hits

    P_GiveArmor = True
End Function

' P_GiveCard
Private Sub P_GiveCard(ByVal player As Long, ByVal card As Long)
    If players(player).cards(card) Then Exit Sub

    players(player).bonuscount = BONUSADD
    players(player).cards(card) = True
End Sub

' P_GivePower
Public Function P_GivePower(ByVal player As Long, ByVal power As Long) As Boolean
    P_GivePower = True

    If power = pw_invulnerability Then
        players(player).powers(power) = INVULNTICS
        Exit Function
    End If

    If power = pw_invisibility Then
        players(player).powers(power) = INVISTICS
        mobjs(players(player).mo).flags = mobjs(players(player).mo).flags Or MF_SHADOW
        Exit Function
    End If

    If power = pw_infrared Then
        players(player).powers(power) = INFRATICS
        Exit Function
    End If

    If power = pw_ironfeet Then
        players(player).powers(power) = IRONTICS
        Exit Function
    End If

    If power = pw_strength Then
        P_GiveBody player, 100
        players(player).powers(power) = 1
        Exit Function
    End If

    If players(player).powers(power) <> 0 Then
        P_GivePower = False                      ' already got it
        Exit Function
    End If

    players(player).powers(power) = 1
End Function

' P_TouchSpecialThing
Public Sub P_TouchSpecialThing(ByVal special As Long, ByVal toucher As Long)
    Dim player As Long, i As Long, delta As Long, sound As Long, dropped As Boolean

    delta = mobjs(special).z - mobjs(toucher).z

    If delta > mobjs(toucher).height Or delta < -8 * FRACUNIT Then
        ' out of reach
        Exit Sub
    End If

    sound = sfx_itemup
    player = mobjs(toucher).player

    ' Dead thing touching.
    ' Can happen with a sliding player corpse.
    If mobjs(toucher).health <= 0 Then Exit Sub

    dropped = ((mobjs(special).flags And MF_DROPPED) <> 0)

    ' Identify by sprite.
    Select Case mobjs(special).sprite
        ' armor
        Case SPR_ARM1
            If Not P_GiveArmor(player, 1) Then Exit Sub
            players(player).message = GOTARMOR

        Case SPR_ARM2
            If Not P_GiveArmor(player, 2) Then Exit Sub
            players(player).message = GOTMEGA

        ' bonus items
        Case SPR_BON1
            players(player).health = players(player).health + 1 ' can go over 100%
            If players(player).health > 200 Then players(player).health = 200
            mobjs(players(player).mo).health = players(player).health
            players(player).message = GOTHTHBONUS

        Case SPR_BON2
            players(player).armorpoints = players(player).armorpoints + 1 ' can go over 100%
            If players(player).armorpoints > 200 Then players(player).armorpoints = 200
            If players(player).armortype = 0 Then players(player).armortype = 1
            players(player).message = GOTARMBONUS

        Case SPR_SOUL
            players(player).health = players(player).health + 100
            If players(player).health > 200 Then players(player).health = 200
            mobjs(players(player).mo).health = players(player).health
            players(player).message = GOTSUPER
            sound = sfx_getpow

        Case SPR_MEGA
            If gamemode <> commercial Then Exit Sub
            players(player).health = 200
            mobjs(players(player).mo).health = players(player).health
            P_GiveArmor player, 2
            players(player).message = GOTMSPHERE
            sound = sfx_getpow

        ' cards
        ' leave cards for everyone
        Case SPR_BKEY
            If Not players(player).cards(it_bluecard) Then players(player).message = GOTBLUECARD
            P_GiveCard player, it_bluecard
            If netgame Then Exit Sub

        Case SPR_YKEY
            If Not players(player).cards(it_yellowcard) Then players(player).message = GOTYELWCARD
            P_GiveCard player, it_yellowcard
            If netgame Then Exit Sub

        Case SPR_RKEY
            If Not players(player).cards(it_redcard) Then players(player).message = GOTREDCARD
            P_GiveCard player, it_redcard
            If netgame Then Exit Sub

        Case SPR_BSKU
            If Not players(player).cards(it_blueskull) Then players(player).message = GOTBLUESKUL
            P_GiveCard player, it_blueskull
            If netgame Then Exit Sub

        Case SPR_YSKU
            If Not players(player).cards(it_yellowskull) Then players(player).message = GOTYELWSKUL
            P_GiveCard player, it_yellowskull
            If netgame Then Exit Sub

        Case SPR_RSKU
            If Not players(player).cards(it_redskull) Then players(player).message = GOTREDSKULL
            P_GiveCard player, it_redskull
            If netgame Then Exit Sub

        ' medikits, heals
        Case SPR_STIM
            If Not P_GiveBody(player, 10) Then Exit Sub
            players(player).message = GOTSTIM

        Case SPR_MEDI
            If Not P_GiveBody(player, 25) Then Exit Sub

            If players(player).health < 25 Then
                players(player).message = GOTMEDINEED
            Else
                players(player).message = GOTMEDIKIT
            End If

        ' power ups
        Case SPR_PINV
            If Not P_GivePower(player, pw_invulnerability) Then Exit Sub
            players(player).message = GOTINVUL
            sound = sfx_getpow

        Case SPR_PSTR
            If Not P_GivePower(player, pw_strength) Then Exit Sub
            players(player).message = GOTBERSERK
            If players(player).readyweapon <> wp_fist Then players(player).pendingweapon = wp_fist
            sound = sfx_getpow

        Case SPR_PINS
            If Not P_GivePower(player, pw_invisibility) Then Exit Sub
            players(player).message = GOTINVIS
            sound = sfx_getpow

        Case SPR_SUIT
            If Not P_GivePower(player, pw_ironfeet) Then Exit Sub
            players(player).message = GOTSUIT
            sound = sfx_getpow

        Case SPR_PMAP
            If Not P_GivePower(player, pw_allmap) Then Exit Sub
            players(player).message = GOTMAP
            sound = sfx_getpow

        Case SPR_PVIS
            If Not P_GivePower(player, pw_infrared) Then Exit Sub
            players(player).message = GOTVISOR
            sound = sfx_getpow

        ' ammo
        Case SPR_CLIP
            If dropped Then
                If Not P_GiveAmmo(player, am_clip, 0) Then Exit Sub
            Else
                If Not P_GiveAmmo(player, am_clip, 1) Then Exit Sub
            End If
            players(player).message = GOTCLIP

        Case SPR_AMMO
            If Not P_GiveAmmo(player, am_clip, 5) Then Exit Sub
            players(player).message = GOTCLIPBOX

        Case SPR_ROCK
            If Not P_GiveAmmo(player, am_misl, 1) Then Exit Sub
            players(player).message = GOTROCKET

        Case SPR_BROK
            If Not P_GiveAmmo(player, am_misl, 5) Then Exit Sub
            players(player).message = GOTROCKBOX

        Case SPR_CELL
            If Not P_GiveAmmo(player, am_cell, 1) Then Exit Sub
            players(player).message = GOTCELL

        Case SPR_CELP
            If Not P_GiveAmmo(player, am_cell, 5) Then Exit Sub
            players(player).message = GOTCELLBOX

        Case SPR_SHEL
            If Not P_GiveAmmo(player, am_shell, 1) Then Exit Sub
            players(player).message = GOTSHELLS

        Case SPR_SBOX
            If Not P_GiveAmmo(player, am_shell, 5) Then Exit Sub
            players(player).message = GOTSHELLBOX

        Case SPR_BPAK
            If Not players(player).backpack Then
                For i = 0 To NUMAMMO - 1
                    players(player).maxammo(i) = players(player).maxammo(i) * 2
                Next
                players(player).backpack = True
            End If
            For i = 0 To NUMAMMO - 1
                P_GiveAmmo player, i, 1
            Next
            players(player).message = GOTBACKPACK

        ' weapons
        Case SPR_BFUG
            If Not P_GiveWeapon(player, wp_bfg, False) Then Exit Sub
            players(player).message = GOTBFG9000
            sound = sfx_wpnup

        Case SPR_MGUN
            If Not P_GiveWeapon(player, wp_chaingun, dropped) Then Exit Sub
            players(player).message = GOTCHAINGUN
            sound = sfx_wpnup

        Case SPR_CSAW
            If Not P_GiveWeapon(player, wp_chainsaw, False) Then Exit Sub
            players(player).message = GOTCHAINSAW
            sound = sfx_wpnup

        Case SPR_LAUN
            If Not P_GiveWeapon(player, wp_missile, False) Then Exit Sub
            players(player).message = GOTLAUNCHER
            sound = sfx_wpnup

        Case SPR_PLAS
            If Not P_GiveWeapon(player, wp_plasma, False) Then Exit Sub
            players(player).message = GOTPLASMA
            sound = sfx_wpnup

        Case SPR_SHOT
            If Not P_GiveWeapon(player, wp_shotgun, dropped) Then Exit Sub
            players(player).message = GOTSHOTGUN
            sound = sfx_wpnup

        Case SPR_SGN2
            If Not P_GiveWeapon(player, wp_supershotgun, dropped) Then Exit Sub
            players(player).message = GOTSHOTGUN2
            sound = sfx_wpnup

        Case Else
            I_Error "P_SpecialThing: Unknown gettable thing"
    End Select

    If (mobjs(special).flags And MF_COUNTITEM) <> 0 Then
        players(player).itemcount = players(player).itemcount + 1
    End If
    P_RemoveMobj special
    players(player).bonuscount = players(player).bonuscount + BONUSADD
    If player = consoleplayer Then S_StartSound 0, sound
End Sub

' KillMobj
Private Sub P_KillMobj(ByVal source As Long, ByVal target As Long)
    Dim item As Long, mo As Long, tp As Long

    mobjs(target).flags = mobjs(target).flags And (Not (MF_SHOOTABLE Or MF_FLOAT Or MF_SKULLFLY))

    If mobjs(target).mtype <> MT_SKULL Then mobjs(target).flags = mobjs(target).flags And (Not MF_NOGRAVITY)

    mobjs(target).flags = mobjs(target).flags Or MF_CORPSE Or MF_DROPOFF
    mobjs(target).height = SAR(mobjs(target).height, 2)

    tp = mobjs(target).player

    If source <> 0 And mobjs(source).player >= 0 Then
        ' count for intermission
        If (mobjs(target).flags And MF_COUNTKILL) <> 0 Then
            players(mobjs(source).player).killcount = players(mobjs(source).player).killcount + 1
        End If

        If tp >= 0 Then
            players(mobjs(source).player).frags(tp) = players(mobjs(source).player).frags(tp) + 1
        End If
    ElseIf Not netgame And (mobjs(target).flags And MF_COUNTKILL) <> 0 Then
        ' count all monster deaths, even those caused by other monsters
        players(0).killcount = players(0).killcount + 1
    End If

    If tp >= 0 Then
        ' count environment kills against you
        If source = 0 Then players(tp).frags(tp) = players(tp).frags(tp) + 1

        mobjs(target).flags = mobjs(target).flags And (Not MF_SOLID)
        players(tp).playerstate = PST_DEAD
        P_DropWeapon tp

        If tp = consoleplayer And automapactive Then
            ' don't die in auto map, switch view prior to dying
            AM_Stop
        End If
    End If

    If mobjs(target).health < -mobjinfo(mobjs(target).info).spawnhealth And _
       mobjinfo(mobjs(target).info).xdeathstate <> 0 Then
        P_SetMobjState target, mobjinfo(mobjs(target).info).xdeathstate
    Else
        P_SetMobjState target, mobjinfo(mobjs(target).info).deathstate
    End If
    mobjs(target).tics = mobjs(target).tics - (P_Random() And 3)

    If mobjs(target).tics < 1 Then mobjs(target).tics = 1

    ' Drop stuff.
    ' This determines the kind of object spawned during the death frame of a
    ' thing.
    Select Case mobjs(target).mtype
        Case MT_WOLFSS, MT_POSSESSED
            item = MT_CLIP
        Case MT_SHOTGUY
            item = MT_SHOTGUN
        Case MT_CHAINGUY
            item = MT_CHAINGUN
        Case Else
            Exit Sub
    End Select

    mo = P_SpawnMobj(mobjs(target).x, mobjs(target).y, ONFLOORZ, item)
    mobjs(mo).flags = mobjs(mo).flags Or MF_DROPPED  ' special versions of items
End Sub

' P_DamageMobj - damages both enemies and players.  "inflictor" is the thing
' that caused the damage (creature or missile, can be 0 (slime, etc)).
' "source" is the thing to target after taking damage (creature or 0).
' Source and inflictor are the same for melee attacks.  Source can be 0 for
' slime, barrel explosions and other environmental stuff.
Public Sub P_DamageMobj(ByVal target As Long, ByVal inflictor As Long, ByVal source As Long, ByVal damage As Long)
    Dim ang As Long, saved As Long, player As Long, thrust As Long, temp As Long
    Dim dothrust As Boolean, nochainsaw As Boolean

    If (mobjs(target).flags And MF_SHOOTABLE) = 0 Then Exit Sub ' shouldn't happen...

    If mobjs(target).health <= 0 Then Exit Sub

    If (mobjs(target).flags And MF_SKULLFLY) <> 0 Then
        mobjs(target).momx = 0
        mobjs(target).momy = 0
        mobjs(target).momz = 0
    End If

    player = mobjs(target).player
    If player >= 0 And gameskill = sk_baby Then damage = SAR(damage, 1) ' take half damage in trainer mode

    ' Some close combat weapons should not inflict thrust and push the
    ' victim out of reach, thus kick away unless using the chainsaw.
    dothrust = False
    If inflictor <> 0 And (mobjs(target).flags And MF_NOCLIP) = 0 Then
        If source = 0 Then
            dothrust = True
        ElseIf mobjs(source).player < 0 Then
            dothrust = True
        ElseIf players(mobjs(source).player).readyweapon <> wp_chainsaw Then
            dothrust = True
        End If
    End If

    If dothrust Then
        ang = R_PointToAngle2(mobjs(inflictor).x, mobjs(inflictor).y, mobjs(target).x, mobjs(target).y)

        thrust = damage * (FRACUNIT \ 8) * 100 \ mobjinfo(mobjs(target).info).mass

        ' make fall forwards sometimes
        If damage < 40 Then
            If damage > mobjs(target).health Then
                If mobjs(target).z - mobjs(inflictor).z > 64 * FRACUNIT Then
                    If (P_Random() And 1) <> 0 Then
                        ang = ang + ANG180
                        thrust = thrust * 4
                    End If
                End If
            End If
        End If

        ang = FINEANGLE(ang)
        mobjs(target).momx = mobjs(target).momx + FixedMul(thrust, finecosine(ang))
        mobjs(target).momy = mobjs(target).momy + FixedMul(thrust, finesine(ang))
    End If

    ' player specific
    If player >= 0 Then
        ' end of game hell hack
        If sectors(subsectors(mobjs(target).subsector).sector).special = 11 And damage >= mobjs(target).health Then
            damage = mobjs(target).health - 1
        End If

        ' Below certain threshold, ignore damage in GOD mode, or with INVUL
        ' power.
        If damage < 1000 And ((players(player).cheats And CF_GODMODE) <> 0 Or _
                              players(player).powers(pw_invulnerability) <> 0) Then
            Exit Sub
        End If

        If players(player).armortype <> 0 Then
            If players(player).armortype = 1 Then
                saved = damage \ 3
            Else
                saved = damage \ 2
            End If

            If players(player).armorpoints <= saved Then
                ' armor is used up
                saved = players(player).armorpoints
                players(player).armortype = 0
            End If
            players(player).armorpoints = players(player).armorpoints - saved
            damage = damage - saved
        End If
        players(player).health = players(player).health - damage ' mirror mobj health here for Dave
        If players(player).health < 0 Then players(player).health = 0

        players(player).attacker = source
        players(player).damagecount = players(player).damagecount + damage ' add damage after armor / invuln

        If players(player).damagecount > 100 Then players(player).damagecount = 100 ' teleport stomp does 10k points...

        If damage < 100 Then temp = damage Else temp = 100
        ' (I_Tactile: no force feedback here)
    End If

    ' do the damage
    mobjs(target).health = mobjs(target).health - damage
    If mobjs(target).health <= 0 Then
        P_KillMobj source, target
        Exit Sub
    End If

    If P_Random() < mobjinfo(mobjs(target).info).painchance And (mobjs(target).flags And MF_SKULLFLY) = 0 Then
        mobjs(target).flags = mobjs(target).flags Or MF_JUSTHIT ' fight back!

        P_SetMobjState target, mobjinfo(mobjs(target).info).painstate
    End If

    mobjs(target).reactiontime = 0               ' we're awake now...

    If (mobjs(target).threshold = 0 Or mobjs(target).mtype = MT_VILE) And source <> 0 Then
        If source <> target And mobjs(source).mtype <> MT_VILE Then
            ' if not intent on another player, chase after this one
            mobjs(target).target = source
            mobjs(target).threshold = BASETHRESHOLD
            If mobjs(target).state = mobjinfo(mobjs(target).info).spawnstate And _
               mobjinfo(mobjs(target).info).seestate <> S_NULL Then
                P_SetMobjState target, mobjinfo(mobjs(target).info).seestate
            End If
        End If
    End If
End Sub
