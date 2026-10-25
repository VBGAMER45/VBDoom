"""
gen_stubs.py - writes the per-cluster interface stub modules into stubs/.

Each stub declares the Public routines a cluster must provide, with the exact
VB6 signatures every other cluster codes against.  A cluster's real modules
replace its stub file in the project.
"""
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "stubs")

HDR = """' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
'
' INTERFACE STUB - every Public routine below must be implemented, with exactly
' this signature, by the real modules of the cluster.  The stub is removed from
' the project when those modules are added.
Option Explicit
"""

L = "ByVal line As Long"

STUBS = {
"stubs_r": ("Cluster A: renderer (r_main, r_bsp, r_segs, r_plane, r_things, r_draw, r_data, r_sky)", """
' r_main
Public Sub R_Init(): End Sub
Public Sub R_SetViewSize(ByVal blocks As Long, ByVal detail As Long): End Sub     ' sets setsizeneeded
Public Sub R_ExecuteSetViewSize(): End Sub                                        ' clears setsizeneeded
Public Sub R_RenderPlayerView(ByVal playernum As Long): End Sub
Public Function R_PointOnSide(ByVal x As Long, ByVal y As Long, ByVal node As Long) As Long: End Function
Public Function R_PointOnSegSide(ByVal x As Long, ByVal y As Long, ByVal seg As Long) As Long: End Function
Public Function R_PointToAngle(ByVal x As Long, ByVal y As Long) As Long: End Function          ' relative to viewx/viewy
Public Function R_PointToAngle2(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long: End Function
Public Function R_PointToDist(ByVal x As Long, ByVal y As Long) As Long: End Function
Public Function R_PointInSubsector(ByVal x As Long, ByVal y As Long) As Long: End Function      ' subsector index
' r_things
Public Sub R_InitSprites(): End Sub                                               ' called by P_Init
' r_data
Public Sub R_InitData(): End Sub
Public Sub R_PrecacheLevel(): End Sub
Public Function R_FlatNumForName(ByVal name As String) As Long: End Function          ' 0-based flat number (lump - firstflat)
Public Function R_TextureNumForName(ByVal name As String) As Long: End Function       ' "-" -> 0, I_Error if missing
Public Function R_CheckTextureNumForName(ByVal name As String) As Long: End Function  ' "-" -> 0, -1 if missing
' r_sky
Public Sub R_InitSkyMap(): End Sub
' r_draw
Public Sub R_FillBackScreen(): End Sub
Public Sub R_DrawViewBorder(): End Sub
Public Sub R_VideoErase(ByVal ofs As Long, ByVal count As Long): End Sub
"""),

"stubs_p": ("Cluster B: play core (p_setup, p_tick, p_mobj, p_maputl, p_map, p_sight, p_user, p_inter)", """
' p_setup  (allocates lines(0 To numlines): lines(numlines) is a scratch "junk" line)
Public Sub P_Init(): End Sub
Public Sub P_SetupLevel(ByVal episode As Long, ByVal map As Long, ByVal playermask As Long, ByVal skill As Long): End Sub
' p_tick
Public Sub P_InitThinkers(): End Sub
Public Function P_NewThinker() As Long: End Function          ' free slot, all per-type records zeroed, not linked
Public Sub P_AddThinker(ByVal th As Long): End Sub
Public Sub P_RemoveThinker(ByVal th As Long): End Sub         ' thinkers(th).func = TF_REMOVED, freed by P_RunThinkers
Public Sub P_Ticker(): End Sub
' p_mobj
Public Function P_SpawnMobj(ByVal x As Long, ByVal y As Long, ByVal z As Long, ByVal mtype As Long) As Long: End Function
Public Sub P_RemoveMobj(ByVal mo As Long): End Sub
Public Function P_SetMobjState(ByVal mo As Long, ByVal state As Long) As Boolean: End Function
Public Sub P_MobjThinker(ByVal mo As Long): End Sub
Public Sub P_SpawnPuff(ByVal x As Long, ByVal y As Long, ByVal z As Long): End Sub
Public Sub P_SpawnBlood(ByVal x As Long, ByVal y As Long, ByVal z As Long, ByVal damage As Long): End Sub
Public Function P_SpawnMissile(ByVal source As Long, ByVal dest As Long, ByVal mtype As Long) As Long: End Function
Public Sub P_SpawnPlayerMissile(ByVal source As Long, ByVal mtype As Long): End Sub
Public Sub P_CheckMissileSpawn(ByVal th As Long): End Sub
Public Sub P_ExplodeMissile(ByVal mo As Long): End Sub
Public Sub P_SpawnPlayer(mthing As mapthing_t): End Sub
Public Sub P_SpawnMapThing(mthing As mapthing_t): End Sub
Public Sub P_RespawnSpecials(): End Sub
' p_maputl
Public Function P_AproxDistance(ByVal dx As Long, ByVal dy As Long) As Long: End Function
Public Function P_PointOnLineSide(ByVal x As Long, ByVal y As Long, ByVal line As Long) As Long: End Function
Public Function P_PointOnDivlineSide(ByVal x As Long, ByVal y As Long, dl As divline_t) As Long: End Function
Public Sub P_MakeDivline(ByVal li As Long, dl As divline_t): End Sub
Public Function P_InterceptVector(v2 As divline_t, v1 As divline_t) As Long: End Function
Public Function P_BoxOnLineSide(tmbox() As Long, ByVal ld As Long) As Long: End Function
Public Sub P_LineOpening(ByVal linedef As Long): End Sub
Public Function P_BlockLinesIterator(ByVal x As Long, ByVal y As Long, ByVal func As Long) As Boolean: End Function   ' func = ITER_xxx
Public Function P_BlockThingsIterator(ByVal x As Long, ByVal y As Long, ByVal func As Long) As Boolean: End Function  ' func = ITER_xxx
Public Function P_PathTraverse(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal flags As Long, ByVal trav As Long) As Boolean: End Function ' trav = TRAV_xxx
Public Sub P_UnsetThingPosition(ByVal thing As Long): End Sub
Public Sub P_SetThingPosition(ByVal thing As Long): End Sub
' p_map
Public Function P_CheckPosition(ByVal thing As Long, ByVal x As Long, ByVal y As Long) As Boolean: End Function
Public Function P_TryMove(ByVal thing As Long, ByVal x As Long, ByVal y As Long) As Boolean: End Function
Public Function P_TeleportMove(ByVal thing As Long, ByVal x As Long, ByVal y As Long) As Boolean: End Function
Public Sub P_SlideMove(ByVal mo As Long): End Sub
Public Sub P_UseLines(ByVal player As Long): End Sub
Public Function P_ChangeSector(ByVal sector As Long, ByVal crunch As Boolean) As Boolean: End Function
Public Function P_AimLineAttack(ByVal t1 As Long, ByVal angle As Long, ByVal distance As Long) As Long: End Function
Public Sub P_LineAttack(ByVal t1 As Long, ByVal angle As Long, ByVal distance As Long, ByVal slope As Long, ByVal damage As Long): End Sub
Public Sub P_RadiusAttack(ByVal spot As Long, ByVal source As Long, ByVal damage As Long): End Sub
' p_sight
Public Function P_CheckSight(ByVal t1 As Long, ByVal t2 As Long) As Boolean: End Function
' p_user
Public Sub P_PlayerThink(ByVal player As Long): End Sub
' p_inter
Public Sub P_TouchSpecialThing(ByVal special As Long, ByVal toucher As Long): End Sub
Public Sub P_DamageMobj(ByVal target As Long, ByVal inflictor As Long, ByVal source As Long, ByVal damage As Long): End Sub
Public Function P_GivePower(ByVal player As Long, ByVal power As Long) As Boolean: End Function
"""),

"stubs_ai": ("Cluster C: monster AI and weapons (p_enemy, p_pspr)", """
' p_enemy
Public Sub P_CallMobjAction(ByVal action As Long, ByVal mo As Long): End Sub   ' dispatch AF_xxx -> A_xxx(mo); called by P_SetMobjState
Public Function PIT_VileCheck(ByVal thing As Long) As Boolean: End Function    ' ITER_VILECHECK callback
Public Sub P_NoiseAlert(ByVal target As Long, ByVal emmiter As Long): End Sub
' p_pspr
Public Sub P_SetupPsprites(ByVal player As Long): End Sub
Public Sub P_MovePsprites(ByVal player As Long): End Sub
Public Sub P_DropWeapon(ByVal player As Long): End Sub
Public Sub P_SetPsprite(ByVal player As Long, ByVal position As Long, ByVal stnum As Long): End Sub
"""),

"stubs_spec": ("Cluster D: specials (p_spec, p_lights, p_switch, p_plats, p_doors, p_ceilng, p_floor, p_telept, p_saveg)", """
' p_spec
Public Sub P_InitPicAnims(): End Sub
Public Sub P_SpawnSpecials(): End Sub
Public Sub P_UpdateSpecials(): End Sub
Public Function P_UseSpecialLine(ByVal thing As Long, ByVal line As Long, ByVal side As Long) As Boolean: End Function
Public Sub P_ShootSpecialLine(ByVal thing As Long, ByVal line As Long): End Sub
Public Sub P_CrossSpecialLine(ByVal linenum As Long, ByVal side As Long, ByVal thing As Long): End Sub
Public Sub P_PlayerInSpecialSector(ByVal player As Long): End Sub
Public Function twoSided(ByVal sector As Long, ByVal line As Long) As Long: End Function
Public Function getSector(ByVal currentSector As Long, ByVal line As Long, ByVal side As Long) As Long: End Function
Public Function getSide(ByVal currentSector As Long, ByVal line As Long, ByVal side As Long) As Long: End Function
Public Function getNextSector(ByVal line As Long, ByVal sec As Long) As Long: End Function    ' -1 = none
Public Function P_FindLowestFloorSurrounding(ByVal sec As Long) As Long: End Function
Public Function P_FindHighestFloorSurrounding(ByVal sec As Long) As Long: End Function
Public Function P_FindNextHighestFloor(ByVal sec As Long, ByVal currentheight As Long) As Long: End Function
Public Function P_FindLowestCeilingSurrounding(ByVal sec As Long) As Long: End Function
Public Function P_FindHighestCeilingSurrounding(ByVal sec As Long) As Long: End Function
Public Function P_FindSectorFromLineTag(ByVal line As Long, ByVal start As Long) As Long: End Function
Public Function P_FindMinSurroundingLight(ByVal sector As Long, ByVal max As Long) As Long: End Function
Public Function EV_DoDonut(ByVal line As Long) As Long: End Function
' p_lights
Public Sub T_FireFlicker(ByVal th As Long): End Sub
Public Sub P_SpawnFireFlicker(ByVal sector As Long): End Sub
Public Sub T_LightFlash(ByVal th As Long): End Sub
Public Sub P_SpawnLightFlash(ByVal sector As Long): End Sub
Public Sub T_StrobeFlash(ByVal th As Long): End Sub
Public Sub P_SpawnStrobeFlash(ByVal sector As Long, ByVal fastOrSlow As Long, ByVal inSync As Long): End Sub
Public Sub EV_StartLightStrobing(ByVal line As Long): End Sub
Public Sub EV_TurnTagLightsOff(ByVal line As Long): End Sub
Public Sub EV_LightTurnOn(ByVal line As Long, ByVal bright As Long): End Sub
Public Sub T_Glow(ByVal th As Long): End Sub
Public Sub P_SpawnGlowingLight(ByVal sector As Long): End Sub
' p_switch
Public Sub P_InitSwitchList(): End Sub
Public Sub P_ChangeSwitchTexture(ByVal line As Long, ByVal useAgain As Long): End Sub
' p_plats
Public Sub T_PlatRaise(ByVal th As Long): End Sub
Public Function EV_DoPlat(ByVal line As Long, ByVal typ As Long, ByVal amount As Long) As Long: End Function
Public Sub P_AddActivePlat(ByVal th As Long): End Sub
Public Sub P_RemoveActivePlat(ByVal th As Long): End Sub
Public Sub EV_StopPlat(ByVal line As Long): End Sub
Public Sub P_ActivateInStasis(ByVal tag As Long): End Sub
' p_doors
Public Sub T_VerticalDoor(ByVal th As Long): End Sub
Public Function EV_DoDoor(ByVal line As Long, ByVal typ As Long) As Long: End Function
Public Function EV_DoLockedDoor(ByVal line As Long, ByVal typ As Long, ByVal thing As Long) As Long: End Function
Public Sub EV_VerticalDoor(ByVal line As Long, ByVal thing As Long): End Sub
Public Sub P_SpawnDoorCloseIn30(ByVal sec As Long): End Sub
Public Sub P_SpawnDoorRaiseIn5Mins(ByVal sec As Long, ByVal secnum As Long): End Sub
' p_ceilng
Public Sub T_MoveCeiling(ByVal th As Long): End Sub
Public Function EV_DoCeiling(ByVal line As Long, ByVal typ As Long) As Long: End Function
Public Sub P_AddActiveCeiling(ByVal th As Long): End Sub
Public Sub P_RemoveActiveCeiling(ByVal th As Long): End Sub
Public Function EV_CeilingCrushStop(ByVal line As Long) As Long: End Function
Public Sub P_ActivateInStasisCeiling(ByVal line As Long): End Sub
' p_floor
Public Function T_MovePlane(ByVal sector As Long, ByVal speed As Long, ByVal dest As Long, ByVal crush As Boolean, ByVal floorOrCeiling As Long, ByVal direction As Long) As Long: End Function
Public Sub T_MoveFloor(ByVal th As Long): End Sub
Public Function EV_DoFloor(ByVal line As Long, ByVal floortype As Long) As Long: End Function
Public Function EV_BuildStairs(ByVal line As Long, ByVal typ As Long) As Long: End Function
' p_telept
Public Function EV_Teleport(ByVal line As Long, ByVal side As Long, ByVal thing As Long) As Long: End Function
' p_saveg  (f = VB file number opened For Binary by g_game; use Put/Get)
Public Sub P_ArchivePlayers(ByVal f As Integer): End Sub
Public Sub P_UnArchivePlayers(ByVal f As Integer): End Sub
Public Sub P_ArchiveWorld(ByVal f As Integer): End Sub
Public Sub P_UnArchiveWorld(ByVal f As Integer): End Sub
Public Sub P_ArchiveThinkers(ByVal f As Integer): End Sub
Public Sub P_UnArchiveThinkers(ByVal f As Integer): End Sub
Public Sub P_ArchiveSpecials(ByVal f As Integer): End Sub
Public Sub P_UnArchiveSpecials(ByVal f As Integer): End Sub
"""),

"stubs_g": ("Cluster E: game logic, menus, heads-up (g_game, m_menu, hu_lib, hu_stuff)", """
' g_game
Public Sub G_DeathMatchSpawnPlayer(ByVal playernum As Long): End Sub
Public Sub G_InitNew(ByVal skill As Long, ByVal episode As Long, ByVal map As Long): End Sub
Public Sub G_DeferedInitNew(ByVal skill As Long, ByVal episode As Long, ByVal map As Long): End Sub
Public Sub G_DeferedPlayDemo(ByVal name As String): End Sub
Public Sub G_LoadGame(ByVal name As String): End Sub
Public Sub G_DoLoadGame(): End Sub
Public Sub G_SaveGame(ByVal slot As Long, ByVal description As String): End Sub
Public Sub G_RecordDemo(ByVal name As String): End Sub
Public Sub G_BeginRecording(): End Sub
Public Sub G_PlayDemo(ByVal name As String): End Sub
Public Sub G_TimeDemo(ByVal name As String): End Sub
Public Function G_CheckDemoStatus() As Boolean: End Function
Public Sub G_ExitLevel(): End Sub
Public Sub G_SecretExitLevel(): End Sub
Public Sub G_WorldDone(): End Sub
Public Sub G_Ticker(): End Sub
Public Function G_Responder(ev As event_t) As Boolean: End Function
Public Sub G_ScreenShot(): End Sub
Public Sub G_BuildTiccmd(cmd As ticcmd_t): End Sub
Public Sub G_PlayerReborn(ByVal player As Long): End Sub
Public Sub G_SetTurbo(ByVal turboscale As Long): End Sub           ' -turbo: scales forwardmove()/sidemove()
' m_menu
Public Function M_Responder(ev As event_t) As Boolean: End Function
Public Sub M_Ticker(): End Sub
Public Sub M_Drawer(): End Sub
Public Sub M_Init(): End Sub
Public Sub M_StartControlPanel(): End Sub
' hu_stuff  (HU_Init fills hu_font() in doomdef)
Public Sub HU_Init(): End Sub
Public Sub HU_Start(): End Sub
Public Function HU_Responder(ev As event_t) As Boolean: End Function
Public Sub HU_Ticker(): End Sub
Public Sub HU_Drawer(): End Sub
Public Function HU_dequeueChatChar() As Long: End Function
Public Sub HU_Erase(): End Sub
"""),

"stubs_ui": ("Cluster F: status bar, automap, intermission, finale, wipes (st_lib, st_stuff, am_map, wi_stuff, f_finale, f_wipe)", """
' st_stuff
Public Function ST_Responder(ev As event_t) As Boolean: End Function
Public Sub ST_Ticker(): End Sub
Public Sub ST_Drawer(ByVal fullscreen As Boolean, ByVal refresh As Boolean): End Sub
Public Sub ST_Start(): End Sub
Public Sub ST_Init(): End Sub
' am_map
Public Function AM_Responder(ev As event_t) As Boolean: End Function
Public Sub AM_Ticker(): End Sub
Public Sub AM_Drawer(): End Sub
Public Sub AM_Start(): End Sub
Public Sub AM_Stop(): End Sub
' wi_stuff
Public Sub WI_Ticker(): End Sub
Public Sub WI_Drawer(): End Sub
Public Sub WI_Start(wbstartstruct As wbstartstruct_t): End Sub
' f_finale
Public Function F_Responder(ev As event_t) As Boolean: End Function
Public Sub F_Ticker(): End Sub
Public Sub F_Drawer(): End Sub
Public Sub F_StartFinale(): End Sub
' f_wipe  (uses screens 2 and 3 as scratch, like the C version)
Public Function wipe_StartScreen(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long: End Function
Public Function wipe_EndScreen(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long: End Function
Public Function wipe_ScreenWipe(ByVal wipeno As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal ticks As Long) As Boolean: wipe_ScreenWipe = True: End Function
"""),

"stubs_snd": ("Cluster G: sound (s_sound, i_sound: waveOut mixer and MUS -> MIDI music)", """
' s_sound   origin: 0 = none, > 0 = mobj slot, < 0 = sector sound origin -(sector + 1)
Public Sub S_Init(ByVal sfxVolume As Long, ByVal musicVolume As Long): End Sub
Public Sub S_Start(): End Sub
Public Sub S_StartSound(ByVal origin As Long, ByVal sfx_id As Long): End Sub
Public Sub S_StartSoundAtVolume(ByVal origin As Long, ByVal sfx_id As Long, ByVal volume As Long): End Sub
Public Sub S_StopSound(ByVal origin As Long): End Sub
Public Sub S_StartMusic(ByVal music_id As Long): End Sub
Public Sub S_ChangeMusic(ByVal music_id As Long, ByVal looping As Boolean): End Sub
Public Sub S_StopMusic(): End Sub
Public Sub S_PauseSound(): End Sub
Public Sub S_ResumeSound(): End Sub
Public Sub S_UpdateSounds(ByVal listener As Long): End Sub
Public Sub S_SetMusicVolume(ByVal volume As Long): End Sub
Public Sub S_SetSfxVolume(ByVal volume As Long): End Sub
' i_sound (Win32)
Public Sub I_InitSound(): End Sub
Public Sub I_ShutdownSound(): End Sub
Public Sub I_UpdateSound(): End Sub           ' once per frame: mix and queue audio buffers
Public Sub I_InitMusic(): End Sub
Public Sub I_ShutdownMusic(): End Sub
Public Sub I_UpdateMusic(): End Sub           ' once per frame: send due MIDI events
"""),
}

if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    for name, (title, body) in STUBS.items():
        text = 'Attribute VB_Name = "%s"\n\' %s\n%s%s' % (name, title, HDR, body)
        with open(os.path.join(OUT, name + ".bas"), "w", newline="\r\n") as f:
            f.write(text)
        print("wrote stubs/%s.bas" % name)
