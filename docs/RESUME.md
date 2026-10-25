# DoomVB: resume notes

Handoff for continuing the port in a fresh session. Read this file first, then
`docs/CONVENTIONS.md` (VB6 pitfalls and the C-to-VB translation rules).

## Goal

A complete Visual Basic 6 port of the GPL DOOM source release
(`ref/linuxdoom-1.10`, GPL v2; the port stays GPL). The game data is not part of
the port: it loads any IWAD. `freedoom1.wad` (BSD-licensed, original art) sits in
the project folder for testing; `doom1.wad` / `doom.wad` / `doom2.wad` work too.

The user chose to port one cluster at a time (no parallel agents) and asked for
playable doors and weapons first. DirectX was considered and not needed: the
renderer is pure software, and the GDI blit is not a bottleneck.

## Environment

* Project: `C:\Users\Owner\desktop\websites\doom`
* VB6: `C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE` (command-line
  `/make` works); resource compiler `VB98\Wizards\RC.EXE`
* Python 3.12 runs the tools (`python -I tools\...`)
* The Bash tool works; PowerShell is used for launching the exe and making PNGs

## Build and test

```
python tools\build.py              build DoomVB.exe (writes DoomVB.vbp from src/ + stubs)
python tools\build.py --release    same, without array bounds checks
python tools\check_names.py        find case-insensitive name clashes in one pass
python tools\debugbuild.py         DoomVB_debug.exe with per-procedure call traces
python tools\gen_tables.py         regenerate info/sounds/strings/items tables + res/tables.bin
python tools\gen_stubs.py          regenerate stubs/stubs_*.bas (cluster interfaces)
```

VB6 reports only the first compile error (file and line in `build.log`). Its
line numbers match the file's lines.

Standard automated test (no window, no sound, no mouse grab):

```
DoomVB.exe -hidden -nosound -nomouse -timedemo demo1            -> fps in doomvb.log
DoomVB.exe -hidden -nosound -nomouse -warp 1 1 -shot 40 a.bmp -quitat 50
```

Convert a BMP to PNG to view it (PowerShell, System.Drawing):
`[System.Drawing.Image]::FromFile(bmp)` then `.Save(png, Png)`; scale 2x with
NearestNeighbor so it is readable.

Extra debug options: `-checklinks` (verifies thing lists every tic),
`-synclog file` (per-tic fingerprint; diff two runs), `-watch slot`
(log one thinker slot), `-config file`, `-scale n`, `-width px`.

Scripted input for hidden runs: `-key <tic>[+<n>] <key>` presses a key at a
gametic and releases it 1 (or n) tics later; repeat it as often as needed.
Keys: a letter or digit, or `esc enter space bksp tab up down left right minus
equals shift ctrl alt pause f1..f12`. Use `-config <copy of default.cfg>` so
menu changes do not end up in the real `default.cfg`. Example (new game from
the title screen, then walk forward for a second):

```
DoomVB.exe -hidden -nosound -nomouse -key 20 esc -key 40 enter -key 60 enter
           -key 80 enter -key 120+35 up -shot 160 a.bmp -quitat 170
```

A `-hidden` run ignores the real keyboard (the invisible window can still own
the keyboard focus, and keys typed on the machine used to leak into tests).
`-exitat <tic>` exits the level as the exit switch would: intermission
(`-warp 1 1 -exitat 40`) or finale (`-warp 1 8 -exitat 30`; E1's text phase
lasts about 1600 tics, E3's bunny scroll ends near tic 3000).

Sound tests without opening any device (so nothing is heard):
`-sndcapture <file.wav>` mixes 1260 samples (one tic at 44100 Hz) per gametic
into a WAV file, also under `-timedemo`; `-muslog <file>` lists every music
event as it would be streamed (absolute ms and the MIDI message in hex, plus
song / play / stop / pause / volume lines). A run with sound on logs
`I_ShutdownSound: N buffers played, M underruns` and the MIDI song position
to doomvb.log; use a scratch config with low `sfx_volume` / `music_volume`.

## Current state (all compiled and run)

| Cluster | C files | State |
|---|---|---|
| Foundation | doomdef/doomstat, m_fixed+tables, m_random, m_bbox, m_argv, m_cheat, m_misc, w_wad, v_video, i_system, i_video (+frmDoom), d_main, generated tables | done |
| A renderer | r_main r_bsp r_segs r_plane r_things r_draw r_data r_sky | done |
| B play core | p_setup p_tick p_mobj p_maputl p_map p_sight p_user p_inter | done |
| C AI + weapons | p_enemy p_pspr | done |
| D specials | p_spec p_doors p_floor p_ceilng p_plats p_lights p_switch p_telept p_saveg | done |
| E game/menu/HUD | g_game m_menu hu_lib hu_stuff | done |
| F screens | st_lib st_stuff am_map wi_stuff f_finale f_wipe | done (DOOM 2 cast call untested: no DOOM 2 IWAD here) |
| G sound | s_sound + Win32 i_sound (waveOut mixer, MUS and MIDI via midiStream) | done (MUS path tested with a synthetic song only: no DOOM IWAD here) |
| Network | d_net + Winsock i_net (linuxdoom's UDP protocol) | done (tested with up to 3 copies on one PC; not yet between two PCs) |

About 29,400 lines of VB. Every C module of linuxdoom is ported; no stubs are
compiled in any more.

Starting a net game (linuxdoom's options): each player gives their player
number and the addresses of all the others, e.g. on 192.168.1.10
`DoomVB.exe -net 1 192.168.1.20 -deathmatch -warp 1 1` and on 192.168.1.20
`DoomVB.exe -net 2 192.168.1.10`. Player 1 sends skill, episode, map,
deathmatch/altdeath, nomonsters and respawn to the others; `-fast`, `-timer`,
`-avg`, `-dup` and `-extratic` are not sent and must be the same everywhere
(as in linuxdoom). UDP port 5029, `-port n` to change it; a host can be
`address:port`. Windows Firewall asks once to let DoomVB.exe receive. While
waiting for the others the window title says so; Esc aborts.

Verified: Freedoom demo1 plays all 7117 tics with combat, doors and pickups.
The optimized and the instrumented build agree tic for tic (`-synclog`).
Rendering: about 950 fps with optimization off and bounds checks on.

Menus and HUD verified with `-key` runs and screenshots: Esc menu (also over
the title demo), new game through episode and skill menus, Nightmare prompt,
options (messages, detail, screen size, mouse), F-keys (F1 help, F2/F3
save/load, F5 detail, F6/F9 quicksave/quickload prompts, F7 end game, F8,
F10 quit with y, F11 gamma), player messages at the top left. Save/load round
trip: the view after loading is pixel-identical to the view at save time.
Savegames are about 3.2 MB (p_saveg writes the arrays whole).

Cluster F verified the same way: status bar (numbers, face states, keys,
arms, cheats iddqd/idkfa, palette flashes in demo1), screen melt, automap
(Tab, grid, iddt twice, zoom, title line), intermission (stats count, Use skips,
"Entering", next level loads), E1 and E3 finales (text, art screen, bunny
scroll). Freedoom's intermission map pieces (WISPLAT, WIURH*, WIA*) are 1x1
blank patches, so no splats or animations show with it. demo1 still matches
the pre-cluster build tic for tic; with the melt and status bar it times at
about 740 fps.

Network verified with two or three hidden copies on this PC (the second and
third run from their own folders in the scratchpad, `p2\DoomVB.exe -iwad ...`,
so each writes its own doomvb.log), each with scripted `-key` input and its
own `-port`, then comparing the `-synclog` files: identical in co-op (2 and 3
players, `-extratic`), deathmatch, `-dup 2`, with `-netdrop 4` (25 % loss,
about 70 resends each way) and over a 1200 tic game with 10 % loss. The
netgame consistency check (players' x every tic) never fired. Screenshots:
chat ("INDIGO: HI" on player 1), FRAGS in the status bar, "Player 2 left the
game" (after the chat message timed out, as in DOOM), co-op and deathmatch
intermissions. Without loss a 2 player game runs at full speed (1200 tics in
37 s, single player 35 s); loss costs speed (each lost packet stalls until the
resend: about 26 tics/s at 10 %). Single player now takes the C path (node 0
rebound packets): demo1 still matches tic for tic, timedemo and real time.

Sound (cluster G) verified: a pistol shot captured with `-sndcapture` matches
a Python model of the mixer sample for sample (22050 Hz lump, linuxdoom's
random pitch, volume, separation); demo1's capture has no stuck channels and
is byte-identical between the release and the instrumented build. The music
event streams of D_E1M6, D_E1M1 and D_INTER (across its loop point) match an
independent Python MIDI parser event for event; a hand-made MUS lump in a
PWAD (`-file`) covering every MUS event type converts as expected. Live runs
with the real devices: waveOut and the MIDI mapper open, no underruns
(after feeding the mixer inside the melt's wait loop too), pause/unpause and
the intermission song change work. demo1 still matches tic for tic (sound
calls M_Random, never P_Random).

## Scaffolding

* `stubs/` and `tools/gen_stubs.py` are no longer used by the build (every
  cluster is complete, so `tools/build.py` drops each stub); they document
  the cluster interfaces.

## Next steps (agreed order)

1. `--release` performance check, play-through tests (E1M1 to the E1M8
   boss), and doom2-format maps (freedoom2.wad). README.md exists (controls,
   multiplayer, options, building, license); the user adds screenshots as
   docs/screenshots/{ingame,automap,menu,deathmatch}.png. Writing it found
   `-avg` broken (`20 * 60 * 35` overflowed 16 bits; now `20&`).
   For GitHub: freedoom1.wad (Freedoom 0.13.0) is committed, with its
   COPYING, CREDITS and CREDITS-MUSIC (from the v0.13.0 tag) in freedoom/;
   .gitignore keeps the exes, logs, default.cfg and savegames out (DoomVB.exe
   goes into a GitHub Release, which must also carry freedoom1.wad,
   freedoom/, LICENSE.TXT and README.md); .gitattributes forces CRLF on the
   VB6 sources.
   (The user is happy with linuxdoom's random sound pitch: no switch wanted.)
2. A net game between two real PCs (the user has to run it).
3. Optional: smaller savegames; a start dialog for net games instead of the
   command line.

## Interfaces and facts the remaining modules need

* Patches are **lump numbers**: `V_DrawPatch x, y, scrn, lump` (also
  `V_DrawPatchDirect`, `V_DrawPatchFlipped`); header via `PatchWidth/Height/
  LeftOffset/TopOffset(lump)`. Lump by name: `W_GetNumForName("STBAR")`.
* Screens: one byte array `screens()`; screen n starts at `n * SCREENSIZE`.
  The status bar draws its background on screen 4 and copies with `V_CopyRect`.
  `f_wipe` uses screens 2 and 3. `I_ReadScreen scrn` copies screen 0 to scrn.
* Palette: `I_SetPalette offset` takes a **wadbuf offset**:
  `W_CacheLumpName("PLAYPAL") + palnum * 768` (st_stuff's ST_doPaletteStuff).
* `hu_font(0 To 62)` (lumps STCFN033..STCFN095) is declared in doomdef and
  filled by HU_Init. `toupper(c)` (doomdef) is C's toupper for font lookups.
* hu_lib widgets (`hu_textline_t`, `hu_stext_t`, `hu_itext_t`) always use
  hu_font; their draw/erase routines take the "on" flag by value (C: a
  boolean pointer). HU_Drawer draws the map title when `automapactive`.
* Menu function pointers are id constants (`MI_`, `MD_`, `MR_` in m_menu).
* st_lib widgets keep a copy of their patch list; their update routines take
  the shown value and the "on" flag by value (C: int* / boolean* pointers).
* The automap posts `AM_MSGENTERED` / `AM_MSGEXITED` (doomdef) to ST_Responder.
* Player message: `players(p).message` is a String, "" = no message.
* Strings from d_englsh.h are Public variables in `dstrings.bas` (`E1TEXT`,
  `GOTARMOR`, `QSPROMPT` ...); `%s`/`%d` placeholders need `Replace`. Character
  constants (`HUSTR_KEYGREEN` ...) are Long consts.
* Cheats: `m_cheat.bas` provides `cht_MakeSeq cht, text` (plain text is
  scrambled for you; `"idclev" & Chr$(1) & Chr$(0) & Chr$(0)` for parameter
  cheats), `cht_CheckCheat(cht, key)`, `cht_GetParam(cht)` (returns String).
* Events: `event_t` (`typ`, data1..3); responders take `ev As event_t` ByRef.
  Key codes are Doom's (KEY_* constants, lowercase ASCII letters).
* Config vars (screenblocks, detailLevel, showMessages, usegamma,
  mouseSensitivity, snd_SfxVolume, snd_MusicVolume, numChannels, chat_macros)
  are globals loaded/saved by `m_misc` (`default.cfg`).
* Savegames: `G_SaveGame slot, description` and `G_LoadGame filename`; files are
  `DoomExePath & "doomsav" & slot & ".dsg"` (SAVEGAMENAME in m_misc). Header is
  24-byte description + "version 110".
* Sound origins: 0 = none, > 0 = mobj slot, < 0 = sector `-(sector + 1)`
  (`SectorSoundOrg(sec)`). Tables `S_sfx()` / `S_music()` (sounds_tables);
  sfx lumps are "DS" & name, music lumps "D_" & name. `I_UpdateSound` and
  `I_UpdateMusic` are called every frame from D_DoomLoop and also inside the
  wait loops of TryRunTics and the screen melt (VB6 has no threads; the mixer
  keeps only 60 ms queued). `I_InitSound` / `I_InitMusic` from I_Init;
  shutdown from I_Quit / I_Error. `-nosound`, `-nosfx`, `-nomusic` honoured.
* i_sound: 8 mixer channels at 44100 Hz, each sound stepped at its own rate
  (Freedoom effects are mostly 22050 Hz). Music: MUS or MIDI lumps become one
  list of millisecond-timed events (a reset prelude, then the song, then a
  NOP at the song length); the midiStream API plays 500 ms chunks from it,
  up to six queued. The music volume scales controller 7 at streaming time.
* The window procedure (`i_video.DoomWndProc`) posts key events; Alt is safe for
  strafing (system menu suppressed); Alt+F4 / window close set `quitRequested`.
* `G_CheckDemoStatus` shows a MsgBox with the timedemo result unless `-hidden`.

## Lessons learned (also in CONVENTIONS.md)

* **VB6 "Optimize for Fast Code" miscompiled EV_Teleport** (the wrong mobj got
  teleported, which later corrupted sector thing lists). The build now uses No
  Optimization. If two builds ever disagree, compare them with `-synclog`.
* C `&&`/`||` short-circuit and VB `And`/`Or` do not: nest Ifs whenever the
  right side indexes an array or calls `P_Random` (demo sync depends on the
  exact P_Random call order).
* 16-bit literal overflow (`10240 * 4`), `&H8000` is negative, `\` vs `/`,
  `>>` on negatives, unsigned angle compares (ULT/UGT), comparisons binding
  tighter than And.
* Reserved names bit us: `scale` (even as a UDT member), `Empty`, `line`, `Me`.
  Case-insensitive clashes: VIEWHEIGHT/viewheight, NUMSPRITES/numsprites, module
  m_random vs M_Random, SP_SECRET/sp_secret inside one module (check_names.py
  only compares across modules). Module-level declarations must precede
  procedures.
* The sources are CRLF. A Python edit script that reads in text mode gets `\n`
  line ends; splitting that on `\r\n` gives one "line", and a line filter then
  empties the file (this destroyed g_game.bas once; it was rebuilt from the
  session transcript). Prefer the Edit tool, and check counts before writing.
* Never `ReDim` the thinker arrays inside a With or ByRef-element call; growth
  happens at the start of P_Ticker.
* Masked mid textures built from several patches are drawn solid
  (`R_DrawSolidMaskedColumn`); vanilla showed the "medusa" garbage.
* Deliberate small deviations from vanilla (all marked in comments): dynamic
  limits (visplanes, drawsegs, openings, vissprites, intercepts, spechit,
  solidsegs), guards against C crashes (EV_DoDonut, A_BrainSpit, zero distance
  in R_PointToDist), demo versions 109 and 110 accepted, retail counts as
  episode 2 for switch textures, savegame format specific to this port.
  Menus/HUD: the quit message comes from the game's own eight (as in the DOS
  executables; linuxdoom's table had two missing commas and crude unused
  lines), the detail toggle works (linuxdoom disabled it), Plutonia/TNT get
  their own map titles, and the French key tables are used only for the
  French version (the C tested the enum constant `french`, always true).
  Screens: the same `if (french)` / `if (commercial)` enum-constant tests are
  fixed in wi_stuff (intermission animations draw; WIOBJ, which Freedoom lacks,
  only for French), idclev works in DOOM 2 (the C set epsd = 0 and then refused
  it), idmus refuses selections below the first song, AM_Stop sends a proper
  AM_MSGEXITED event, and f_finale's font test uses >= HU_FONTSIZE.
  Sound: effects use the DOS volume scale (s_sound keeps snd_SfxVolume * 8;
  linuxdoom dropped the * 8, making effects 8 times quieter and the map 8
  falloff wrong); I_StopSound, I_SoundIsPlaying and I_UpdateSoundParams work
  as in DOS (linuxdoom's were dummies, so sounds of one origin overlapped
  instead of cutting each other off), the pitch of a playing sound is not
  updated; linuxdoom's random pitch shift (up to about 3 semitones, as in
  DOOM 1.2) is kept; a missing effect plays dspistol (linuxdoom's getsfx) and
  a linked one its link's lump; a missing song plays nothing instead of
  I_Error; music exists at all (linuxdoom's was a dummy) and also reads MIDI
  files; left/right volumes are clamped instead of I_Error.
  Network: packets are sent from the listening socket and a node is matched
  by address and port (linuxdoom: a second socket, address only), so several
  copies can play on one PC; hosts may be `address:port`; when every host is
  127.x the socket listens on 127.0.0.1 only; a "connection reset" from a
  host that is not listening yet is ignored; ticcmd fields are wrapped to the
  C types (char/short/byte) before use, so the local copy equals what the
  others receive; packets with a bad player number or tic count are dropped
  (the C wrote past its arrays); the player number and number of hosts are
  checked; `-debugfile` writes debug<player>.txt to the exe folder; while
  waiting for the start the window is up and its title says so (no console
  on Windows); the wait loops sleep and feed the sound; gametime starts at
  the current time (the C began at 0, so the first tics piled up).
