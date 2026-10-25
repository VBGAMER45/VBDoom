Attribute VB_Name = "doomdef"
' doomdef.bas - shared constants, enums, data structures and global state.
' Combines doomdef.h, doomdata.h, doomstat.h, d_event.h, d_ticcmd.h, d_player.h,
' d_think.h, r_defs.h, r_state.h, p_mobj.h, p_spec.h, p_local.h and info.h (types).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.
' This program is free software; you can redistribute it and/or modify it under
' the terms of the GNU General Public License version 2 (see LICENSE.TXT).
'
' POINTER CONVENTIONS (see docs/CONVENTIONS.md):
'   * Map data arrays (vertexes, sectors, sides, lines, segs, subsectors, nodes)
'     are 0-based; a reference to them is a Long index and -1 means NULL.
'   * Thinkers (mobjs, ceilings, doors, ...) share ONE 1-based index space
'     (the thinker slot); a reference is a Long and 0 means NULL.
'   * Players are 0-based (0..MAXPLAYERS-1); -1 means NULL.
Option Explicit

' ------------------------------------------------------------------ Win32
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ------------------------------------------------------------------ doomdef.h
Public Const VERSION As Long = 110

Public Enum GameMode_t
    shareware = 0
    registered = 1
    commercial = 2
    retail = 3
    indetermined = 4
End Enum

Public Enum GameMission_t
    doom = 0
    doom2 = 1
    pack_tnt = 2
    pack_plut = 3
    mission_none = 4
End Enum

Public Enum Language_t
    english = 0
    french = 1
    german = 2
    lang_unknown = 3
End Enum

Public Const SCREENWIDTH As Long = 320
Public Const SCREENHEIGHT As Long = 200
Public Const SCREENSIZE As Long = 64000          ' SCREENWIDTH * SCREENHEIGHT
Public Const MAXPLAYERS As Long = 4
Public Const TICRATE As Long = 35

Public Enum gamestate_t
    GS_LEVEL = 0
    GS_INTERMISSION = 1
    GS_FINALE = 2
    GS_DEMOSCREEN = 3
    GS_FORCEWIPE = -1                            ' wipegamestate = -1 forces a wipe
End Enum

Public Const MTF_EASY As Long = 1
Public Const MTF_NORMAL As Long = 2
Public Const MTF_HARD As Long = 4
Public Const MTF_AMBUSH As Long = 8

Public Enum skill_t
    sk_baby = 0
    sk_easy = 1
    sk_medium = 2
    sk_hard = 3
    sk_nightmare = 4
End Enum

Public Enum card_t
    it_bluecard = 0
    it_yellowcard = 1
    it_redcard = 2
    it_blueskull = 3
    it_yellowskull = 4
    it_redskull = 5
End Enum
Public Const NUMCARDS As Long = 6

Public Enum weapontype_t
    wp_fist = 0
    wp_pistol = 1
    wp_shotgun = 2
    wp_chaingun = 3
    wp_missile = 4
    wp_plasma = 5
    wp_bfg = 6
    wp_chainsaw = 7
    wp_supershotgun = 8
    wp_nochange = 10
End Enum
Public Const NUMWEAPONS As Long = 9

Public Enum ammotype_t
    am_clip = 0
    am_shell = 1
    am_cell = 2
    am_misl = 3
    am_noammo = 5
End Enum
Public Const NUMAMMO As Long = 4

Public Enum powertype_t
    pw_invulnerability = 0
    pw_strength = 1
    pw_invisibility = 2
    pw_ironfeet = 3
    pw_allmap = 4
    pw_infrared = 5
End Enum
Public Const NUMPOWERS As Long = 6

Public Const INVULNTICS As Long = 30 * TICRATE
Public Const INVISTICS As Long = 60 * TICRATE
Public Const INFRATICS As Long = 120 * TICRATE
Public Const IRONTICS As Long = 60 * TICRATE

' Doom key codes (event data1).  Letters/digits are lowercase ASCII.
Public Const KEY_RIGHTARROW As Long = &HAE
Public Const KEY_LEFTARROW As Long = &HAC
Public Const KEY_UPARROW As Long = &HAD
Public Const KEY_DOWNARROW As Long = &HAF
Public Const KEY_ESCAPE As Long = 27
Public Const KEY_ENTER As Long = 13
Public Const KEY_TAB As Long = 9
Public Const KEY_F1 As Long = &HBB
Public Const KEY_F2 As Long = &HBC
Public Const KEY_F3 As Long = &HBD
Public Const KEY_F4 As Long = &HBE
Public Const KEY_F5 As Long = &HBF
Public Const KEY_F6 As Long = &HC0
Public Const KEY_F7 As Long = &HC1
Public Const KEY_F8 As Long = &HC2
Public Const KEY_F9 As Long = &HC3
Public Const KEY_F10 As Long = &HC4
Public Const KEY_F11 As Long = &HD7
Public Const KEY_F12 As Long = &HD8
Public Const KEY_BACKSPACE As Long = 127
Public Const KEY_PAUSE As Long = &HFF
Public Const KEY_EQUALS As Long = &H3D
Public Const KEY_MINUS As Long = &H2D
Public Const KEY_RSHIFT As Long = &HB6
Public Const KEY_RCTRL As Long = &H9D
Public Const KEY_RALT As Long = &HB8
Public Const KEY_LALT As Long = &HB8

' ------------------------------------------------------------------ m_fixed.h / tables.h
Public Const FRACBITS As Long = 16
Public Const FRACUNIT As Long = 65536
Public Const MAXINT As Long = &H7FFFFFFF
Public Const MININT As Long = &H80000000

Public Const FINEANGLES As Long = 8192
Public Const FINEMASK As Long = 8191
Public Const ANGLETOFINESHIFT As Long = 19
Public Const ANG45 As Long = &H20000000
Public Const ANG90 As Long = &H40000000
Public Const ANG180 As Long = &H80000000
Public Const ANG270 As Long = &HC0000000
Public Const SLOPERANGE As Long = 2048
Public Const SLOPEBITS As Long = 11
Public Const DBITS As Long = 5                   ' FRACBITS - SLOPEBITS

' ------------------------------------------------------------------ doomdata.h
Public Const ML_LABEL As Long = 0
Public Const ML_THINGS As Long = 1
Public Const ML_LINEDEFS As Long = 2
Public Const ML_SIDEDEFS As Long = 3
Public Const ML_VERTEXES As Long = 4
Public Const ML_SEGS As Long = 5
Public Const ML_SSECTORS As Long = 6
Public Const ML_NODES As Long = 7
Public Const ML_SECTORS As Long = 8
Public Const ML_REJECT As Long = 9
Public Const ML_BLOCKMAP As Long = 10

Public Const ML_BLOCKING As Long = 1
Public Const ML_BLOCKMONSTERS As Long = 2
Public Const ML_TWOSIDED As Long = 4
Public Const ML_DONTPEGTOP As Long = 8
Public Const ML_DONTPEGBOTTOM As Long = 16
Public Const ML_SECRET As Long = 32
Public Const ML_SOUNDBLOCK As Long = 64
Public Const ML_DONTDRAW As Long = 128
Public Const ML_MAPPED As Long = 256

Public Const NF_SUBSECTOR As Long = &H8000&

Public Type mapthing_t
    x As Long
    y As Long
    angle As Long
    typ As Long                                  ' C: type
    options As Long
End Type

' ------------------------------------------------------------------ d_event.h / d_ticcmd.h
Public Enum evtype_t
    ev_keydown = 0
    ev_keyup = 1
    ev_mouse = 2
    ev_joystick = 3
End Enum

Public Type event_t
    typ As Long                                  ' evtype_t
    data1 As Long
    data2 As Long
    data3 As Long
End Type

Public Enum gameaction_t
    ga_nothing = 0
    ga_loadlevel = 1
    ga_newgame = 2
    ga_loadgame = 3
    ga_savegame = 4
    ga_playdemo = 5
    ga_completed = 6
    ga_victory = 7
    ga_worlddone = 8
    ga_screenshot = 9
End Enum

Public Const BT_ATTACK As Long = 1
Public Const BT_USE As Long = 2
Public Const BT_SPECIAL As Long = 128
Public Const BT_SPECIALMASK As Long = 3
Public Const BT_CHANGE As Long = 4
Public Const BT_WEAPONMASK As Long = 56
Public Const BT_WEAPONSHIFT As Long = 3
Public Const BTS_PAUSE As Long = 1
Public Const BTS_SAVEGAME As Long = 2
Public Const BTS_SAVEMASK As Long = 28
Public Const BTS_SAVESHIFT As Long = 2

Public Const MAXEVENTS As Long = 64

' forwardmove/sidemove are signed chars (-128..127); angleturn/consistancy are
' signed shorts (-32768..32767); chatchar/buttons are bytes (0..255).
' Code that builds a ticcmd must keep the values inside those ranges.
Public Type ticcmd_t
    forwardmove As Long
    sidemove As Long
    angleturn As Long
    consistancy As Long
    chatchar As Long
    buttons As Long
End Type

' ------------------------------------------------------------------ info.h (types)
Public Const FF_FULLBRIGHT As Long = &H8000&
Public Const FF_FRAMEMASK As Long = &H7FFF&

Public Type state_t
    sprite As Long                               ' spritenum_t
    frame As Long
    tics As Long
    action As Long                               ' actionf_t (AF_xxx)
    nextstate As Long                            ' statenum_t
    misc1 As Long
    misc2 As Long
End Type

Public Type mobjinfo_t
    doomednum As Long
    spawnstate As Long
    spawnhealth As Long
    seestate As Long
    seesound As Long
    reactiontime As Long
    attacksound As Long
    painstate As Long
    painchance As Long
    painsound As Long
    meleestate As Long
    missilestate As Long
    deathstate As Long
    xdeathstate As Long
    deathsound As Long
    speed As Long
    radius As Long
    height As Long
    mass As Long
    damage As Long
    activesound As Long
    flags As Long
    raisestate As Long
End Type

' ------------------------------------------------------------------ sounds.h (types)
Public Type sfxinfo_t
    name As String
    singularity As Boolean
    priority As Long
    link As Long                                 ' sfx index or -1
    pitch As Long
    volume As Long
    lumpnum As Long
    usefulness As Long
End Type

Public Type musicinfo_t
    name As String
    lumpnum As Long
    handle As Long
End Type

' ------------------------------------------------------------------ d_items.h
Public Type weaponinfo_t
    ammo As Long
    upstate As Long
    downstate As Long
    readystate As Long
    atkstate As Long
    flashstate As Long
End Type

' ------------------------------------------------------------------ p_mobj.h
Public Const MF_SPECIAL As Long = 1
Public Const MF_SOLID As Long = 2
Public Const MF_SHOOTABLE As Long = 4
Public Const MF_NOSECTOR As Long = 8
Public Const MF_NOBLOCKMAP As Long = 16
Public Const MF_AMBUSH As Long = 32
Public Const MF_JUSTHIT As Long = 64
Public Const MF_JUSTATTACKED As Long = 128
Public Const MF_SPAWNCEILING As Long = 256
Public Const MF_NOGRAVITY As Long = 512
Public Const MF_DROPOFF As Long = &H400&
Public Const MF_PICKUP As Long = &H800&
Public Const MF_NOCLIP As Long = &H1000&
Public Const MF_SLIDE As Long = &H2000&
Public Const MF_FLOAT As Long = &H4000&
Public Const MF_TELEPORT As Long = &H8000&
Public Const MF_MISSILE As Long = &H10000
Public Const MF_DROPPED As Long = &H20000
Public Const MF_SHADOW As Long = &H40000
Public Const MF_NOBLOOD As Long = &H80000
Public Const MF_CORPSE As Long = &H100000
Public Const MF_INFLOAT As Long = &H200000
Public Const MF_COUNTKILL As Long = &H400000
Public Const MF_COUNTITEM As Long = &H800000
Public Const MF_SKULLFLY As Long = &H1000000
Public Const MF_NOTDMATCH As Long = &H2000000
Public Const MF_TRANSLATION As Long = &HC000000
Public Const MF_TRANSSHIFT As Long = 26

Public Type mobj_t
    x As Long
    y As Long
    z As Long
    snext As Long                                ' sector thing list (mobj, 0 = none)
    sprev As Long
    angle As Long                                ' angle_t (unsigned semantics!)
    sprite As Long
    frame As Long
    bnext As Long                                ' blockmap list (mobj, 0 = none)
    bprev As Long
    subsector As Long                            ' subsector index
    floorz As Long
    ceilingz As Long
    radius As Long
    height As Long
    momx As Long
    momy As Long
    momz As Long
    validcount As Long
    mtype As Long                                ' C: type (mobjtype_t)
    info As Long                                 ' C: info pointer -> index into mobjinfo()
    tics As Long
    state As Long                                ' statenum_t (C: state pointer)
    flags As Long
    health As Long
    movedir As Long
    movecount As Long
    target As Long                               ' mobj, 0 = none
    reactiontime As Long
    threshold As Long
    player As Long                               ' player number, -1 = none
    lastlook As Long
    spawnpoint As mapthing_t
    tracer As Long                               ' mobj, 0 = none
End Type

' ------------------------------------------------------------------ d_think.h
' thinker function ids (C: thinker.function.acp1)
Public Const TF_REMOVED As Long = -1             ' C: function == (actionf_v)(-1)
Public Const TF_NONE As Long = 0                 ' C: function == NULL (in stasis)
Public Const TF_MOBJ As Long = 1                 ' P_MobjThinker
Public Const TF_CEILING As Long = 2              ' T_MoveCeiling
Public Const TF_DOOR As Long = 3                 ' T_VerticalDoor
Public Const TF_FLOOR As Long = 4                ' T_MoveFloor
Public Const TF_PLAT As Long = 5                 ' T_PlatRaise
Public Const TF_FIREFLICKER As Long = 6          ' T_FireFlicker
Public Const TF_LIGHTFLASH As Long = 7           ' T_LightFlash
Public Const TF_STROBE As Long = 8               ' T_StrobeFlash
Public Const TF_GLOW As Long = 9                 ' T_Glow

Public Type thinker_t
    prv As Long
    nxt As Long
    func As Long                                 ' TF_xxx
End Type

' ------------------------------------------------------------------ p_pspr.h
Public Const ps_weapon As Long = 0
Public Const ps_flash As Long = 1
Public Const NUMPSPRITES As Long = 2

Public Type pspdef_t
    state As Long                                ' statenum_t, S_NULL (0) = no sprite
    tics As Long
    sx As Long
    sy As Long
End Type

' ------------------------------------------------------------------ d_player.h
Public Enum playerstate_t
    PST_LIVE = 0
    PST_DEAD = 1
    PST_REBORN = 2
End Enum

Public Const CF_NOCLIP As Long = 1
Public Const CF_GODMODE As Long = 2
Public Const CF_NOMOMENTUM As Long = 4

Public Type player_t
    mo As Long                                   ' mobj, 0 = none
    playerstate As Long
    cmd As ticcmd_t
    viewz As Long
    viewheight As Long
    deltaviewheight As Long
    bob As Long
    health As Long
    armorpoints As Long
    armortype As Long
    powers(0 To 5) As Long                       ' NUMPOWERS
    cards(0 To 5) As Boolean                     ' NUMCARDS
    backpack As Boolean
    frags(0 To 3) As Long                        ' MAXPLAYERS
    readyweapon As Long
    pendingweapon As Long
    weaponowned(0 To 8) As Boolean               ' NUMWEAPONS
    ammo(0 To 3) As Long                         ' NUMAMMO
    maxammo(0 To 3) As Long
    attackdown As Boolean
    usedown As Boolean
    cheats As Long
    refire As Long
    killcount As Long
    itemcount As Long
    secretcount As Long
    message As String                            ' "" = no message (C: NULL)
    damagecount As Long
    bonuscount As Long
    attacker As Long                             ' mobj, 0 = none
    extralight As Long
    fixedcolormap As Long
    colormap As Long
    psprites(0 To 1) As pspdef_t                 ' NUMPSPRITES
    didsecret As Boolean
End Type

Public Type wbplayerstruct_t
    inGame As Boolean                            ' C: in
    skills As Long
    sitems As Long
    ssecret As Long
    stime As Long
    frags(0 To 3) As Long
    score As Long
End Type

Public Type wbstartstruct_t
    epsd As Long
    didsecret As Boolean
    last As Long
    nxt As Long                                  ' C: next
    maxkills As Long
    maxitems As Long
    maxsecret As Long
    maxfrags As Long
    partime As Long
    pnum As Long
    plyr(0 To 3) As wbplayerstruct_t
End Type

' ------------------------------------------------------------------ r_defs.h
Public Const SIL_NONE As Long = 0
Public Const SIL_BOTTOM As Long = 1
Public Const SIL_TOP As Long = 2
Public Const SIL_BOTH As Long = 3

Public Type vertex_t
    x As Long
    y As Long
End Type

Public Type sector_t
    floorheight As Long
    ceilingheight As Long
    floorpic As Long
    ceilingpic As Long
    lightlevel As Long
    special As Long
    tag As Long
    soundtraversed As Long
    soundtarget As Long                          ' mobj, 0 = none
    blockbox(0 To 3) As Long
    soundorgx As Long                            ' C: soundorg (degenmobj_t)
    soundorgy As Long
    soundorgz As Long
    validcount As Long
    thinglist As Long                            ' mobj, 0 = none
    specialdata As Long                          ' thinker, 0 = none
    linecount As Long
    lines As Long                                ' index of first entry in linebuffer()
End Type

Public Type side_t
    textureoffset As Long
    rowoffset As Long
    toptexture As Long
    bottomtexture As Long
    midtexture As Long
    sector As Long
End Type

Public Const ST_HORIZONTAL As Long = 0
Public Const ST_VERTICAL As Long = 1
Public Const ST_POSITIVE As Long = 2
Public Const ST_NEGATIVE As Long = 3

Public Type line_t
    v1 As Long                                   ' vertex index
    v2 As Long
    dx As Long
    dy As Long
    flags As Long
    special As Long
    tag As Long
    sidenum(0 To 1) As Long                      ' side index, -1 = none
    bbox(0 To 3) As Long
    slopetype As Long
    frontsector As Long                          ' sector index, -1 = none
    backsector As Long
    validcount As Long
    specialdata As Long                          ' thinker, 0 = none
End Type

Public Type subsector_t
    sector As Long
    numlines As Long
    firstline As Long
End Type

Public Type seg_t
    v1 As Long
    v2 As Long
    offset As Long
    angle As Long
    sidedef As Long
    linedef As Long
    frontsector As Long
    backsector As Long                           ' -1 = none
End Type

Public Type node_t
    x As Long
    y As Long
    dx As Long
    dy As Long
    bbox(0 To 1, 0 To 3) As Long
    children(0 To 1) As Long                     ' 0..65535, NF_SUBSECTOR flag
End Type

' ------------------------------------------------------------------ m_bbox.h
Public Const BOXTOP As Long = 0
Public Const BOXBOTTOM As Long = 1
Public Const BOXLEFT As Long = 2
Public Const BOXRIGHT As Long = 3

' ------------------------------------------------------------------ p_local.h
Public Const FLOATSPEED As Long = 4 * FRACUNIT
Public Const MAXHEALTH As Long = 100
Public Const PLAYER_VIEWHEIGHT As Long = 41 * FRACUNIT    ' C: VIEWHEIGHT (clashes with viewheight in VB)
Public Const MAPBLOCKUNITS As Long = 128
Public Const MAPBLOCKSIZE As Long = 128 * FRACUNIT
Public Const MAPBLOCKSHIFT As Long = 23          ' FRACBITS + 7
Public Const MAPBMASK As Long = 8388607          ' MAPBLOCKSIZE - 1
Public Const MAPBTOFRAC As Long = 7              ' MAPBLOCKSHIFT - FRACBITS
Public Const PLAYERRADIUS As Long = 16 * FRACUNIT
Public Const MAXRADIUS As Long = 32 * FRACUNIT
Public Const GRAVITY As Long = FRACUNIT
Public Const MAXMOVE As Long = 30 * FRACUNIT
Public Const USERANGE As Long = 64 * FRACUNIT
Public Const MELEERANGE As Long = 64 * FRACUNIT
Public Const MISSILERANGE As Long = 32 * 64 * FRACUNIT
Public Const BASETHRESHOLD As Long = 100
Public Const ONFLOORZ As Long = MININT
Public Const ONCEILINGZ As Long = MAXINT
Public Const ITEMQUESIZE As Long = 128
Public Const MAXINTERCEPTS As Long = 128
Public Const PT_ADDLINES As Long = 1
Public Const PT_ADDTHINGS As Long = 2
Public Const PT_EARLYOUT As Long = 4

Public Type divline_t
    x As Long
    y As Long
    dx As Long
    dy As Long
End Type

Public Type intercept_t
    frac As Long
    isaline As Boolean
    dthing As Long                               ' C: d.thing  (mobj, when Not isaline)
    dline As Long                                ' C: d.line   (line index, when isaline)
End Type

' Iterator / traverser callback ids.  C passes function pointers to
' P_BlockLinesIterator, P_BlockThingsIterator and P_PathTraverse; this port
' passes one of these ids and p_maputl dispatches to the PIT_/PTR_ routine.
Public Const ITER_CHECKLINE As Long = 1          ' PIT_CheckLine          (p_map)
Public Const ITER_ADDLINEINTERCEPTS As Long = 2  ' PIT_AddLineIntercepts  (p_maputl)
Public Const ITER_CHECKTHING As Long = 11        ' PIT_CheckThing         (p_map)
Public Const ITER_STOMPTHING As Long = 12        ' PIT_StompThing         (p_map)
Public Const ITER_CHANGESECTOR As Long = 13      ' PIT_ChangeSector       (p_map)
Public Const ITER_RADIUSATTACK As Long = 14      ' PIT_RadiusAttack       (p_map)
Public Const ITER_VILECHECK As Long = 15         ' PIT_VileCheck          (p_enemy)
Public Const ITER_ADDTHINGINTERCEPTS As Long = 16 ' PIT_AddThingIntercepts (p_maputl)
Public Const TRAV_SLIDE As Long = 21             ' PTR_SlideTraverse      (p_map)
Public Const TRAV_AIM As Long = 22               ' PTR_AimTraverse        (p_map)
Public Const TRAV_SHOOT As Long = 23             ' PTR_ShootTraverse      (p_map)
Public Const TRAV_USE As Long = 24               ' PTR_UseTraverse        (p_map)

' ------------------------------------------------------------------ p_spec.h
Public Const GLOWSPEED As Long = 8
Public Const STROBEBRIGHT As Long = 5
Public Const FASTDARK As Long = 15
Public Const SLOWDARK As Long = 35
Public Const MAXSWITCHES As Long = 50
Public Const MAXBUTTONS As Long = 16
Public Const BUTTONTIME As Long = 35
Public Const PLATWAIT As Long = 3
Public Const PLATSPEED As Long = FRACUNIT
Public Const MAXPLATS As Long = 30
Public Const VDOORSPEED As Long = FRACUNIT * 2
Public Const VDOORWAIT As Long = 150
Public Const CEILSPEED As Long = FRACUNIT
Public Const CEILWAIT As Long = 150
Public Const MAXCEILINGS As Long = 30
Public Const FLOORSPEED As Long = FRACUNIT

Public Type fireflicker_t
    sector As Long
    count As Long
    maxlight As Long
    minlight As Long
End Type

Public Type lightflash_t
    sector As Long
    count As Long
    maxlight As Long
    minlight As Long
    maxtime As Long
    mintime As Long
End Type

Public Type strobe_t
    sector As Long
    count As Long
    minlight As Long
    maxlight As Long
    darktime As Long
    brighttime As Long
End Type

Public Type glow_t
    sector As Long
    minlight As Long
    maxlight As Long
    direction As Long
End Type

' bwhere_e
Public Const bw_top As Long = 0
Public Const bw_middle As Long = 1
Public Const bw_bottom As Long = 2

Public Type button_t
    linenum As Long                              ' C: line  (line index, -1 = none)
    where As Long
    btexture As Long
    btimer As Long
    soundorg As Long                             ' sound origin code (see S_StartSound)
End Type

' plat_e
Public Const plat_up As Long = 0
Public Const plat_down As Long = 1
Public Const plat_waiting As Long = 2
Public Const plat_in_stasis As Long = 3

' plattype_e
Public Const perpetualRaise As Long = 0
Public Const downWaitUpStay As Long = 1
Public Const raiseAndChange As Long = 2
Public Const raiseToNearestAndChange As Long = 3
Public Const blazeDWUS As Long = 4

Public Type plat_t
    sector As Long
    speed As Long
    low As Long
    high As Long
    wait As Long
    count As Long
    status As Long
    oldstatus As Long
    crush As Boolean
    tag As Long
    typ As Long                                  ' plattype_e
End Type

' vldoor_e
Public Const vld_normal As Long = 0
Public Const vld_close30ThenOpen As Long = 1
Public Const vld_close As Long = 2
Public Const vld_open As Long = 3
Public Const vld_raiseIn5Mins As Long = 4
Public Const vld_blazeRaise As Long = 5
Public Const vld_blazeOpen As Long = 6
Public Const vld_blazeClose As Long = 7

Public Type vldoor_t
    typ As Long                                  ' vldoor_e
    sector As Long
    topheight As Long
    speed As Long
    direction As Long
    topwait As Long
    topcountdown As Long
End Type

' ceiling_e
Public Const lowerToFloor As Long = 0
Public Const raiseToHighest As Long = 1
Public Const lowerAndCrush As Long = 2
Public Const crushAndRaise As Long = 3
Public Const fastCrushAndRaise As Long = 4
Public Const silentCrushAndRaise As Long = 5

Public Type ceiling_t
    typ As Long                                  ' ceiling_e
    sector As Long
    bottomheight As Long
    topheight As Long
    speed As Long
    crush As Boolean
    direction As Long
    tag As Long
    olddirection As Long
End Type

' floor_e
Public Const lowerFloor As Long = 0
Public Const lowerFloorToLowest As Long = 1
Public Const turboLower As Long = 2
Public Const raiseFloor As Long = 3
Public Const raiseFloorToNearest As Long = 4
Public Const raiseToTexture As Long = 5
Public Const lowerAndChange As Long = 6
Public Const raiseFloor24 As Long = 7
Public Const raiseFloor24AndChange As Long = 8
Public Const raiseFloorCrush As Long = 9
Public Const raiseFloorTurbo As Long = 10
Public Const donutRaise As Long = 11
Public Const raiseFloor512 As Long = 12

' stair_e
Public Const build8 As Long = 0
Public Const turbo16 As Long = 1

Public Type floormove_t
    typ As Long                                  ' floor_e
    crush As Boolean
    sector As Long
    direction As Long
    newspecial As Long
    texture As Long
    floordestheight As Long
    speed As Long
End Type

' result_e
Public Const res_ok As Long = 0
Public Const res_crushed As Long = 1
Public Const res_pastdest As Long = 2

' ------------------------------------------------------------------ doomstat.h
Public nomonsters As Boolean
Public respawnparm As Boolean
Public fastparm As Boolean
Public devparm As Boolean
Public gamemode As Long                          ' GameMode_t
Public gamemission As Long                       ' GameMission_t
Public language As Long
Public modifiedgame As Boolean

Public startskill As Long
Public startepisode As Long
Public startmap As Long
Public autostart As Boolean

Public gameskill As Long
Public gameepisode As Long
Public gamemap As Long
Public respawnmonsters As Boolean
Public netgame As Boolean
Public deathmatch As Long                        ' 0, 1, or 2 (altdeath)

Public snd_SfxVolume As Long                     ' 0..15 (C: snd_SfxVolume)
Public snd_MusicVolume As Long                   ' 0..15

Public statusbaractive As Boolean
Public automapactive As Boolean
Public menuactive As Boolean
Public paused As Boolean
Public viewactive As Boolean
Public nodrawers As Boolean
Public noblit As Boolean

Public viewwindowx As Long
Public viewwindowy As Long
Public viewheight As Long
Public viewwidth As Long
Public scaledviewwidth As Long
Public viewangleoffset As Long

Public consoleplayer As Long
Public displayplayer As Long

Public totalkills As Long
Public totalitems As Long
Public totalsecret As Long
Public levelstarttic As Long
Public leveltime As Long

Public usergame As Boolean
Public demoplayback As Boolean
Public demorecording As Boolean
Public singledemo As Boolean
Public timingdemo As Boolean

Public gamestate As Long                         ' gamestate_t
Public gametic As Long
Public gameaction As Long                        ' gameaction_t

Public players(0 To 3) As player_t
Public playeringame(0 To 3) As Boolean

Public Const MAX_DM_STARTS As Long = 10
Public deathmatchstarts(0 To 9) As mapthing_t
Public deathmatch_p As Long                      ' number of entries used in deathmatchstarts()
Public playerstarts(0 To 3) As mapthing_t

Public wminfo As wbstartstruct_t

Public maxammo(0 To 3) As Long                   ' set in P_Init/p_inter
Public clipammo(0 To 3) As Long

Public basedefault As String
Public precache As Boolean
Public wipegamestate As Long
Public mouseSensitivity As Long
Public bodyqueslot As Long
Public skyflatnum As Long

Public inhelpscreens As Boolean
Public advancedemo As Boolean
Public setsizeneeded As Boolean
Public sendpause As Boolean
Public sendsave As Boolean
Public screenblocks As Long                      ' m_menu: 3..11
Public detailLevel As Long                       ' m_menu: 0 high, 1 low
Public showMessages As Long                      ' m_menu: 0/1
Public usegamma As Long                          ' 0..4
Public message_dontfuckwithme As Boolean
Public chat_on As Boolean
Public player_names(0 To 3) As String

' key bindings (m_misc defaults)
Public key_right As Long
Public key_left As Long
Public key_up As Long
Public key_down As Long
Public key_strafeleft As Long
Public key_straferight As Long
Public key_fire As Long
Public key_use As Long
Public key_strafe As Long
Public key_speed As Long
Public mousebfire As Long
Public mousebstrafe As Long
Public mousebforward As Long
Public usemouse As Long

' events (d_main)
Public events(0 To 63) As event_t
Public eventhead As Long
Public eventtail As Long

' ------------------------------------------------------------------ r_state.h (map data)
Public numvertexes As Long
Public vertexes() As vertex_t
Public numsegs As Long
Public segs() As seg_t
Public numsectors As Long
Public sectors() As sector_t
Public numsubsectors As Long
Public subsectors() As subsector_t
Public numnodes As Long
Public nodes() As node_t
Public numlines As Long
Public lines() As line_t
Public numsides As Long
Public sides() As side_t
Public linebuffer() As Long                      ' sector line lists (sector_t.lines indexes here)

' blockmap / reject (p_setup)
Public blockmaplump() As Long                    ' whole BLOCKMAP lump as signed shorts
Public blockmap As Long                          ' offset in blockmaplump() of the block offsets (= 4)
Public bmapwidth As Long
Public bmapheight As Long
Public bmaporgx As Long
Public bmaporgy As Long
Public blocklinks() As Long                      ' mobj chains per block, 0 = none
Public rejectmatrix() As Byte

' r_state.h (renderer data shared with the play code)
Public textureheight() As Long
Public spritewidth() As Long
Public spriteoffset() As Long
Public spritetopoffset() As Long
Public colormaps() As Byte                       ' 34 * 256 bytes from the COLORMAP lump
Public firstflat As Long
Public lastflat As Long
Public numflats As Long
Public flattranslation() As Long
Public numtextures As Long
Public texturetranslation() As Long
Public firstspritelump As Long
Public lastspritelump As Long
Public numspritelumps As Long

Public viewx As Long
Public viewy As Long
Public viewz As Long
Public viewangle As Long
Public viewplayer As Long                        ' player number
Public validcount As Long
Public skytexture As Long
Public skytexturemid As Long
Public extralight As Long
Public detailshift As Long

' ------------------------------------------------------------------ thinkers / mobjs (p_tick)
' All thinkers live in one 1-based slot space.  thinkers(0) is thinkercap.
' The per-type arrays are indexed by the same slot number.
Public thinkers() As thinker_t
Public mobjs() As mobj_t
Public ceilings() As ceiling_t
Public doors() As vldoor_t
Public floormoves() As floormove_t
Public plats() As plat_t
Public fireflickers() As fireflicker_t
Public lightflashes() As lightflash_t
Public strobes() As strobe_t
Public glows() As glow_t
Public maxthinkers As Long                       ' current UBound of the thinker arrays

' p_local.h
Public itemrespawnque(0 To 127) As mapthing_t
Public itemrespawntime(0 To 127) As Long
Public iquehead As Long
Public iquetail As Long

Public intercepts() As intercept_t                ' grown by p_maputl (C: MAXINTERCEPTS = 128)
Public intercept_p As Long                       ' number of intercepts used
Public trace As divline_t

Public opentop As Long
Public openbottom As Long
Public openrange As Long
Public lowfloor As Long

Public floatok As Boolean
Public tmfloorz As Long
Public tmceilingz As Long
Public ceilingline As Long                       ' line index, -1 = none
Public linetarget As Long                        ' mobj, 0 = none
Public attackrange As Long
Public topslope As Long
Public bottomslope As Long

' p_spec.h
Public levelTimer As Boolean
Public levelTimeCount As Long
Public buttonlist(0 To 15) As button_t
Public activeplats(0 To 29) As Long              ' plat thinker slot, 0 = none
Public activeceilings(0 To 29) As Long           ' ceiling thinker slot, 0 = none

' ------------------------------------------------------------------ f_wipe.h
Public Const wipe_ColorXForm As Long = 0
Public Const wipe_Melt As Long = 1

' ------------------------------------------------------------------ hu_stuff.h
Public Const HU_FONTSTART As Long = 33           ' the first font characters ('!')
Public Const HU_FONTEND As Long = 95             ' the last font characters ('_')
Public Const HU_FONTSIZE As Long = 63
Public hu_font(0 To 62) As Long                  ' patch lumps, loaded by HU_Init

' ------------------------------------------------------------------ am_map.h
' Used by ST StatusBar stuff: the automap posts these as ev_keyup events.
Public Const AM_MSGHEADER As Long = &H616D0000      ' ('a' << 24) + ('m' << 16)
Public Const AM_MSGENTERED As Long = &H616D6500     ' AM_MSGHEADER | ('e' << 8)
Public Const AM_MSGEXITED As Long = &H616D7800      ' AM_MSGHEADER | ('x' << 8)

' ------------------------------------------------------------------ r_sky.h
Public Const SKYFLATNAME As String = "F_SKY1"
Public Const ANGLETOSKYSHIFT As Long = 22

' ------------------------------------------------------------------ v_video.h
' screens 0..4 (320x200 each) are stored back to back in one array:
' pixel (x, y) of screen n is screens(n * SCREENSIZE + y * SCREENWIDTH + x)
Public screens() As Byte
Public dirtybox(0 To 3) As Long
Public gammatable(0 To 4, 0 To 255) As Byte

' ------------------------------------------------------------------ misc
Public wadfilenames() As String
Public numwadfiles As Long
Public DoomExePath As String

' ------------------------------------------------------------------ angle / unsigned helpers
' (kept here because they are used by every cluster; the hot paths inline them)

' angle >> ANGLETOFINESHIFT for an unsigned angle_t held in a Long
Public Function FINEANGLE(ByVal a As Long) As Long
    FINEANGLE = ((a And &HFFF80000) \ &H80000) And &H1FFF&
End Function

' unsigned comparisons for angle_t / unsigned values held in Longs
Public Function ULT(ByVal a As Long, ByVal b As Long) As Boolean
    ULT = (a Xor &H80000000) < (b Xor &H80000000)
End Function

Public Function UGT(ByVal a As Long, ByVal b As Long) As Boolean
    UGT = (a Xor &H80000000) > (b Xor &H80000000)
End Function

Public Function ULE(ByVal a As Long, ByVal b As Long) As Boolean
    ULE = (a Xor &H80000000) <= (b Xor &H80000000)
End Function

Public Function UGE(ByVal a As Long, ByVal b As Long) As Boolean
    UGE = (a Xor &H80000000) >= (b Xor &H80000000)
End Function

' C toupper() in the "C" locale: only a..z change
Public Function toupper(ByVal c As Long) As Long
    If c >= 97 And c <= 122 Then toupper = c - 32 Else toupper = c
End Function

' sound origin codes for S_StartSound: 0 = none, > 0 = mobj slot,
' < 0 = sector sound origin (-(sector + 1)).
Public Function SectorSoundOrg(ByVal sec As Long) As Long
    SectorSoundOrg = -(sec + 1)
End Function
