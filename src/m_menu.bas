Attribute VB_Name = "m_menu"
' m_menu.bas - DOOM selection menu, options, episode etc.  Sliders and
' icons.  Kinda widget stuff (m_menu.c).
'
' The menus are records in menus(), addressed by the C variable names
' (MainDef, EpiDef, ...).  The C function pointers become id constants:
' MI_xxx for item routines, MD_xxx for draw routines and MR_xxx for message
' responses, run by the dispatchers at the end of this module.
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const SAVESTRINGSIZE As Long = 24
Private Const SKULLXOFF As Long = -32
Private Const LINEHEIGHT As Long = 16

' MENU TYPEDEFS
Private Type menuitem_t
    status As Long                               ' 0 = no cursor here, 1 = ok, 2 = arrows ok, -1 = blank line
    name As String                               ' patch lump, "" = none
    routine As Long                              ' MI_xxx: choice = menu item #; if status = 2, choice=0:leftarrow,1:rightarrow
    alphaKey As Long                             ' hotkey in menu
End Type

Private Type menu_t
    numitems As Long                             ' # of menu items
    prevMenu As Long                             ' previous menu, -1 = none
    menuitems(0 To 7) As menuitem_t              ' menu items
    routine As Long                              ' MD_xxx draw routine
    x As Long
    y As Long                                    ' x,y of menu
    lastOn As Long                               ' last item user was on in menu
End Type

' menus
Private Const MainDef As Long = 0
Private Const EpiDef As Long = 1
Private Const NewDef As Long = 2
Private Const OptionsDef As Long = 3
Private Const ReadDef1 As Long = 4
Private Const ReadDef2 As Long = 5
Private Const SoundDef As Long = 6
Private Const LoadDef As Long = 7
Private Const SaveDef As Long = 8

' main_e
Private Const main_readthis As Long = 4
Private Const main_quitdoom As Long = 5
Private Const main_end As Long = 6
' episodes_e
Private Const ep_end As Long = 4
' newgame_e
Private Const newg_end As Long = 5
' options_e
Private Const opt_messages As Long = 1
Private Const opt_detail As Long = 2
Private Const opt_scrnsize As Long = 3
Private Const opt_mousesens As Long = 5
Private Const opt_end As Long = 8
' sound_e
Private Const sfx_vol As Long = 0
Private Const music_vol As Long = 2
Private Const sound_end As Long = 4
' load_e
Private Const load_end As Long = 6

' item routines (C: void (*routine)(int choice))
Private Const MI_NONE As Long = 0
Private Const MI_NewGame As Long = 1
Private Const MI_Options As Long = 2
Private Const MI_LoadGame As Long = 3
Private Const MI_SaveGame As Long = 4
Private Const MI_ReadThis As Long = 5
Private Const MI_QuitDOOM As Long = 6
Private Const MI_Episode As Long = 7
Private Const MI_ChooseSkill As Long = 8
Private Const MI_EndGame As Long = 9
Private Const MI_ChangeMessages As Long = 10
Private Const MI_ChangeDetail As Long = 11
Private Const MI_SizeDisplay As Long = 12
Private Const MI_ChangeSensitivity As Long = 13
Private Const MI_Sound As Long = 14
Private Const MI_ReadThis2 As Long = 15
Private Const MI_FinishReadThis As Long = 16
Private Const MI_SfxVol As Long = 17
Private Const MI_MusicVol As Long = 18
Private Const MI_LoadSelect As Long = 19
Private Const MI_SaveSelect As Long = 20

' draw routines (C: void (*routine)())
Private Const MD_NONE As Long = 0
Private Const MD_DrawMainMenu As Long = 1
Private Const MD_DrawEpisode As Long = 2
Private Const MD_DrawNewGame As Long = 3
Private Const MD_DrawOptions As Long = 4
Private Const MD_DrawReadThis1 As Long = 5
Private Const MD_DrawReadThis2 As Long = 6
Private Const MD_DrawSound As Long = 7
Private Const MD_DrawLoad As Long = 8
Private Const MD_DrawSave As Long = 9

' message responses (C: void (*messageRoutine)(int response))
Private Const MR_NONE As Long = 0
Private Const MR_QuickSaveResponse As Long = 1
Private Const MR_QuickLoadResponse As Long = 2
Private Const MR_VerifyNightmare As Long = 3
Private Const MR_EndGameResponse As Long = 4
Private Const MR_QuitResponse As Long = 5

Private menus(0 To 8) As menu_t

' temp for screenblocks (0-9)
Private screenSize As Long

' -1 = no quicksave slot picked!
Private quickSaveSlot As Long

' 1 = message to be printed
Private messageToPrint As Long
' ...and here is the message string!
Private messageString As String

Private messageLastMenuActive As Boolean

' timed message = no input from user
Private messageNeedsInput As Boolean

Private messageRoutine As Long                   ' MR_xxx

Private gammamsg(0 To 4) As String

' we are going to be entering a savegame string
Private saveStringEnter As Long
Private saveSlot As Long                         ' which slot to save in
Private saveCharIndex As Long                    ' which char we're editing
' old save description before edit
Private saveOldString As String

Private savegamestrings(0 To 9) As String

Private endstring As String

Private itemOn As Long                           ' menu item skull is on
Private skullAnimCounter As Long                 ' skull animation counter
Private whichSkull As Long                       ' which skull to draw

' graphic name of skulls
Private skullName(0 To 1) As String

' current menudef
Private currentMenu As Long

Private epi As Long

Private detailNames(0 To 1) As String
Private msgNames(0 To 1) As String

Private quitsounds(0 To 7) As Long
Private quitsounds2(0 To 7) As Long

' Deviation: linuxdoom's endmsg[] lost two commas (gluing pairs of messages
' together) and M_QuitDOOM picked from all of it, including crude lines no
' released game showed.  As in the DOS executables, DOOM shows QUITMSG or one
' of its seven messages and DOOM 2 QUITMSG or one of its seven.
Private endmsg(0 To 14) As String

' M_Responder statics
Private joywait As Long
Private mousewait As Long
Private mousey As Long
Private lasty As Long
Private mousex As Long
Private lastx As Long

Private Sub SetMenu(ByVal m As Long, ByVal numitems As Long, ByVal prevMenu As Long, _
                    ByVal routine As Long, ByVal x As Long, ByVal y As Long, ByVal lastOn As Long)
    menus(m).numitems = numitems
    menus(m).prevMenu = prevMenu
    menus(m).routine = routine
    menus(m).x = x
    menus(m).y = y
    menus(m).lastOn = lastOn
End Sub

Private Sub SetItem(ByVal m As Long, ByVal i As Long, ByVal status As Long, ByVal nm As String, _
                    ByVal routine As Long, ByVal key As String)
    menus(m).menuitems(i).status = status
    menus(m).menuitems(i).name = nm
    menus(m).menuitems(i).routine = routine
    If key = "" Then
        menus(m).menuitems(i).alphaKey = 0
    Else
        menus(m).menuitems(i).alphaKey = Asc(key)
    End If
End Sub

' The C initialisers of the menus and tables.
Private Sub InitMenus()
    Dim i As Long

    ' DOOM MENU
    SetMenu MainDef, main_end, -1, MD_DrawMainMenu, 97, 64, 0
    SetItem MainDef, 0, 1, "M_NGAME", MI_NewGame, "n"
    SetItem MainDef, 1, 1, "M_OPTION", MI_Options, "o"
    SetItem MainDef, 2, 1, "M_LOADG", MI_LoadGame, "l"
    SetItem MainDef, 3, 1, "M_SAVEG", MI_SaveGame, "s"
    ' Another hickup with Special edition.
    SetItem MainDef, 4, 1, "M_RDTHIS", MI_ReadThis, "r"
    SetItem MainDef, 5, 1, "M_QUITG", MI_QuitDOOM, "q"

    ' EPISODE SELECT
    SetMenu EpiDef, ep_end, MainDef, MD_DrawEpisode, 48, 63, 0
    SetItem EpiDef, 0, 1, "M_EPI1", MI_Episode, "k"
    SetItem EpiDef, 1, 1, "M_EPI2", MI_Episode, "t"
    SetItem EpiDef, 2, 1, "M_EPI3", MI_Episode, "i"
    SetItem EpiDef, 3, 1, "M_EPI4", MI_Episode, "t"

    ' NEW GAME
    SetMenu NewDef, newg_end, EpiDef, MD_DrawNewGame, 48, 63, sk_medium  ' hurtme
    SetItem NewDef, 0, 1, "M_JKILL", MI_ChooseSkill, "i"
    SetItem NewDef, 1, 1, "M_ROUGH", MI_ChooseSkill, "h"
    SetItem NewDef, 2, 1, "M_HURT", MI_ChooseSkill, "h"
    SetItem NewDef, 3, 1, "M_ULTRA", MI_ChooseSkill, "u"
    SetItem NewDef, 4, 1, "M_NMARE", MI_ChooseSkill, "n"

    ' OPTIONS MENU
    SetMenu OptionsDef, opt_end, MainDef, MD_DrawOptions, 60, 37, 0
    SetItem OptionsDef, 0, 1, "M_ENDGAM", MI_EndGame, "e"
    SetItem OptionsDef, 1, 1, "M_MESSG", MI_ChangeMessages, "m"
    SetItem OptionsDef, 2, 1, "M_DETAIL", MI_ChangeDetail, "g"
    SetItem OptionsDef, 3, 2, "M_SCRNSZ", MI_SizeDisplay, "s"
    SetItem OptionsDef, 4, -1, "", MI_NONE, ""
    SetItem OptionsDef, 5, 2, "M_MSENS", MI_ChangeSensitivity, "m"
    SetItem OptionsDef, 6, -1, "", MI_NONE, ""
    SetItem OptionsDef, 7, 1, "M_SVOL", MI_Sound, "s"

    ' Read This! MENU 1 & 2
    SetMenu ReadDef1, 1, MainDef, MD_DrawReadThis1, 280, 185, 0
    SetItem ReadDef1, 0, 1, "", MI_ReadThis2, ""
    SetMenu ReadDef2, 1, ReadDef1, MD_DrawReadThis2, 330, 175, 0
    SetItem ReadDef2, 0, 1, "", MI_FinishReadThis, ""

    ' SOUND VOLUME MENU
    SetMenu SoundDef, sound_end, OptionsDef, MD_DrawSound, 80, 64, 0
    SetItem SoundDef, 0, 2, "M_SFXVOL", MI_SfxVol, "s"
    SetItem SoundDef, 1, -1, "", MI_NONE, ""
    SetItem SoundDef, 2, 2, "M_MUSVOL", MI_MusicVol, "m"
    SetItem SoundDef, 3, -1, "", MI_NONE, ""

    ' LOAD GAME MENU and SAVE GAME MENU
    SetMenu LoadDef, load_end, MainDef, MD_DrawLoad, 80, 54, 0
    SetMenu SaveDef, load_end, MainDef, MD_DrawSave, 80, 54, 0
    For i = 0 To load_end - 1
        SetItem LoadDef, i, 1, "", MI_LoadSelect, CStr(i + 1)
        SetItem SaveDef, i, 1, "", MI_SaveSelect, CStr(i + 1)
    Next

    gammamsg(0) = GAMMALVL0
    gammamsg(1) = GAMMALVL1
    gammamsg(2) = GAMMALVL2
    gammamsg(3) = GAMMALVL3
    gammamsg(4) = GAMMALVL4

    skullName(0) = "M_SKULL1"
    skullName(1) = "M_SKULL2"
    detailNames(0) = "M_GDHIGH"
    detailNames(1) = "M_GDLOW"
    msgNames(0) = "M_MSGOFF"
    msgNames(1) = "M_MSGON"

    quitsounds(0) = sfx_pldeth: quitsounds(1) = sfx_dmpain
    quitsounds(2) = sfx_popain: quitsounds(3) = sfx_slop
    quitsounds(4) = sfx_telept: quitsounds(5) = sfx_posit1
    quitsounds(6) = sfx_posit3: quitsounds(7) = sfx_sgtatk

    quitsounds2(0) = sfx_vilact: quitsounds2(1) = sfx_getpow
    quitsounds2(2) = sfx_boscub: quitsounds2(3) = sfx_slop
    quitsounds2(4) = sfx_skeswg: quitsounds2(5) = sfx_kntdth
    quitsounds2(6) = sfx_bspact: quitsounds2(7) = sfx_sgtatk

    endmsg(0) = QUITMSG
    ' DOOM1
    endmsg(1) = "please don't leave, there's more" & vbLf & "demons to toast!"
    endmsg(2) = "let's beat it -- this is turning" & vbLf & "into a bloodbath!"
    endmsg(3) = "i wouldn't leave if i were you." & vbLf & "dos is much worse."
    endmsg(4) = "you're trying to say you like dos" & vbLf & "better than me, right?"
    endmsg(5) = "don't leave yet -- there's a" & vbLf & "demon around that corner!"
    endmsg(6) = "ya know, next time you come in here" & vbLf & "i'm gonna toast ya."
    endmsg(7) = "go ahead and leave. see if i care."
    ' QuitDOOM II messages
    endmsg(8) = "you want to quit?" & vbLf & "then, thou hast lost an eighth!"
    endmsg(9) = "don't go now, there's a " & vbLf & "dimensional shambler waiting" & vbLf & "at the dos prompt!"
    endmsg(10) = "get outta here and go back" & vbLf & "to your boring programs."
    endmsg(11) = "if i were your boss, i'd " & vbLf & " deathmatch ya in a minute!"
    endmsg(12) = "look, bud. you leave now" & vbLf & "and you forfeit your body count!"
    endmsg(13) = "just leave. when you come" & vbLf & "back, i'll be waiting with a bat."
    endmsg(14) = "you're lucky i don't smack" & vbLf & "you for thinking about leaving."
End Sub

' savegame file of a slot (C: SAVEGAMENAME"%d.dsg"; -cdrom is not supported,
' as in g_game)
Private Function SaveFileName(ByVal slot As Long) As String
    SaveFileName = DoomExePath & SAVEGAMENAME & slot & ".dsg"
End Function

' M_ReadSaveStrings - read the strings from the savegame files
Private Sub M_ReadSaveStrings()
    Dim f As Integer, i As Long, j As Long, nm As String, buf(0 To SAVESTRINGSIZE - 1) As Byte

    For i = 0 To load_end - 1
        nm = SaveFileName(i)
        If Dir$(nm) = "" Then
            savegamestrings(i) = EMPTYSTRING
            menus(LoadDef).menuitems(i).status = 0
        Else
            For j = 0 To SAVESTRINGSIZE - 1
                buf(j) = 0
            Next
            f = FreeFile
            Open nm For Binary Access Read As #f
            Get #f, 1, buf
            Close #f
            savegamestrings(i) = ""
            For j = 0 To SAVESTRINGSIZE - 1
                If buf(j) = 0 Then Exit For
                savegamestrings(i) = savegamestrings(i) & Chr$(buf(j))
            Next
            menus(LoadDef).menuitems(i).status = 1
        End If
    Next
End Sub

' M_LoadGame & Cie.
Private Sub M_DrawLoad()
    Dim i As Long
    V_DrawPatchDirect 72, 28, 0, W_GetNumForName("M_LOADG")
    For i = 0 To load_end - 1
        M_DrawSaveLoadBorder menus(LoadDef).x, menus(LoadDef).y + LINEHEIGHT * i
        M_WriteText menus(LoadDef).x, menus(LoadDef).y + LINEHEIGHT * i, savegamestrings(i)
    Next
End Sub

' Draw border for the savegame description
Private Sub M_DrawSaveLoadBorder(ByVal x As Long, ByVal y As Long)
    Dim i As Long
    V_DrawPatchDirect x - 8, y + 7, 0, W_GetNumForName("M_LSLEFT")
    For i = 0 To 23
        V_DrawPatchDirect x, y + 7, 0, W_GetNumForName("M_LSCNTR")
        x = x + 8
    Next
    V_DrawPatchDirect x, y + 7, 0, W_GetNumForName("M_LSRGHT")
End Sub

' User wants to load this game
Private Sub M_LoadSelect(ByVal choice As Long)
    G_LoadGame SaveFileName(choice)
    M_ClearMenus
End Sub

' Selected from DOOM menu
Private Sub M_LoadGame(ByVal choice As Long)
    If netgame Then
        M_StartMessage LOADNET, MR_NONE, False
        Exit Sub
    End If
    M_SetupNextMenu LoadDef
    M_ReadSaveStrings
End Sub

' M_SaveGame & Cie.
Private Sub M_DrawSave()
    Dim i As Long
    V_DrawPatchDirect 72, 28, 0, W_GetNumForName("M_SAVEG")
    For i = 0 To load_end - 1
        M_DrawSaveLoadBorder menus(LoadDef).x, menus(LoadDef).y + LINEHEIGHT * i
        M_WriteText menus(LoadDef).x, menus(LoadDef).y + LINEHEIGHT * i, savegamestrings(i)
    Next

    If saveStringEnter <> 0 Then
        i = M_StringWidth(savegamestrings(saveSlot))
        M_WriteText menus(LoadDef).x + i, menus(LoadDef).y + LINEHEIGHT * saveSlot, "_"
    End If
End Sub

' M_Responder calls this when user is finished
Private Sub M_DoSave(ByVal slot As Long)
    G_SaveGame slot, savegamestrings(slot)
    M_ClearMenus

    ' PICK QUICKSAVE SLOT YET?
    If quickSaveSlot = -2 Then quickSaveSlot = slot
End Sub

' User wants to save. Start string input for M_Responder
Private Sub M_SaveSelect(ByVal choice As Long)
    ' we are going to be intercepting all chars
    saveStringEnter = 1

    saveSlot = choice
    saveOldString = savegamestrings(choice)
    If savegamestrings(choice) = EMPTYSTRING Then savegamestrings(choice) = ""
    saveCharIndex = Len(savegamestrings(choice))
End Sub

' Selected from DOOM menu
Private Sub M_SaveGame(ByVal choice As Long)
    If Not usergame Then
        M_StartMessage SAVEDEAD, MR_NONE, False
        Exit Sub
    End If

    If gamestate <> GS_LEVEL Then Exit Sub

    M_SetupNextMenu SaveDef
    M_ReadSaveStrings
End Sub

' M_QuickSave
Private Sub M_QuickSaveResponse(ByVal ch As Long)
    If ch = 121 Then                             ' 'y'
        M_DoSave quickSaveSlot
        S_StartSound 0, sfx_swtchx
    End If
End Sub

Private Sub M_QuickSave()
    If Not usergame Then
        S_StartSound 0, sfx_oof
        Exit Sub
    End If

    If gamestate <> GS_LEVEL Then Exit Sub

    If quickSaveSlot < 0 Then
        M_StartControlPanel
        M_ReadSaveStrings
        M_SetupNextMenu SaveDef
        quickSaveSlot = -2                       ' means to pick a slot now
        Exit Sub
    End If
    M_StartMessage Replace(QSPROMPT, "%s", savegamestrings(quickSaveSlot)), MR_QuickSaveResponse, True
End Sub

' M_QuickLoad
Private Sub M_QuickLoadResponse(ByVal ch As Long)
    If ch = 121 Then                             ' 'y'
        M_LoadSelect quickSaveSlot
        S_StartSound 0, sfx_swtchx
    End If
End Sub

Private Sub M_QuickLoad()
    If netgame Then
        M_StartMessage QLOADNET, MR_NONE, False
        Exit Sub
    End If

    If quickSaveSlot < 0 Then
        M_StartMessage QSAVESPOT, MR_NONE, False
        Exit Sub
    End If
    M_StartMessage Replace(QLPROMPT, "%s", savegamestrings(quickSaveSlot)), MR_QuickLoadResponse, True
End Sub

' Read This Menus
' Had a "quick hack to fix romero bug"
Private Sub M_DrawReadThis1()
    inhelpscreens = True
    Select Case gamemode
        Case commercial
            V_DrawPatchDirect 0, 0, 0, W_GetNumForName("HELP")
        Case shareware, registered, retail
            V_DrawPatchDirect 0, 0, 0, W_GetNumForName("HELP1")
    End Select
End Sub

' Read This Menus - optional second page.
Private Sub M_DrawReadThis2()
    inhelpscreens = True
    Select Case gamemode
        Case retail, commercial
            ' This hack keeps us from having to change menus.
            V_DrawPatchDirect 0, 0, 0, W_GetNumForName("CREDIT")
        Case shareware, registered
            V_DrawPatchDirect 0, 0, 0, W_GetNumForName("HELP2")
    End Select
End Sub

' Change Sfx & Music volumes
Private Sub M_DrawSound()
    V_DrawPatchDirect 60, 38, 0, W_GetNumForName("M_SVOL")

    M_DrawThermo menus(SoundDef).x, menus(SoundDef).y + LINEHEIGHT * (sfx_vol + 1), 16, snd_SfxVolume

    M_DrawThermo menus(SoundDef).x, menus(SoundDef).y + LINEHEIGHT * (music_vol + 1), 16, snd_MusicVolume
End Sub

Private Sub M_Sound(ByVal choice As Long)
    M_SetupNextMenu SoundDef
End Sub

Private Sub M_SfxVol(ByVal choice As Long)
    Select Case choice
        Case 0
            If snd_SfxVolume <> 0 Then snd_SfxVolume = snd_SfxVolume - 1
        Case 1
            If snd_SfxVolume < 15 Then snd_SfxVolume = snd_SfxVolume + 1
    End Select

    S_SetSfxVolume snd_SfxVolume                 ' *8
End Sub

Private Sub M_MusicVol(ByVal choice As Long)
    Select Case choice
        Case 0
            If snd_MusicVolume <> 0 Then snd_MusicVolume = snd_MusicVolume - 1
        Case 1
            If snd_MusicVolume < 15 Then snd_MusicVolume = snd_MusicVolume + 1
    End Select

    S_SetMusicVolume snd_MusicVolume             ' *8
End Sub

' M_DrawMainMenu
Private Sub M_DrawMainMenu()
    V_DrawPatchDirect 94, 2, 0, W_GetNumForName("M_DOOM")
End Sub

' M_NewGame
Private Sub M_DrawNewGame()
    V_DrawPatchDirect 96, 14, 0, W_GetNumForName("M_NEWG")
    V_DrawPatchDirect 54, 38, 0, W_GetNumForName("M_SKILL")
End Sub

Private Sub M_NewGame(ByVal choice As Long)
    If netgame And Not demoplayback Then
        M_StartMessage NEWGAME, MR_NONE, False
        Exit Sub
    End If

    If gamemode = commercial Then
        M_SetupNextMenu NewDef
    Else
        M_SetupNextMenu EpiDef
    End If
End Sub

' M_Episode
Private Sub M_DrawEpisode()
    V_DrawPatchDirect 54, 38, 0, W_GetNumForName("M_EPISOD")
End Sub

Private Sub M_VerifyNightmare(ByVal ch As Long)
    If ch <> 121 Then Exit Sub                   ' 'y'

    G_DeferedInitNew sk_nightmare, epi + 1, 1
    M_ClearMenus
End Sub

Private Sub M_ChooseSkill(ByVal choice As Long)
    If choice = sk_nightmare Then
        M_StartMessage NIGHTMARE, MR_VerifyNightmare, True
        Exit Sub
    End If

    G_DeferedInitNew choice, epi + 1, 1
    M_ClearMenus
End Sub

Private Sub M_Episode(ByVal choice As Long)
    If gamemode = shareware And choice <> 0 Then
        M_StartMessage SWSTRING, MR_NONE, False
        M_SetupNextMenu ReadDef1
        Exit Sub
    End If

    ' Yet another hack...
    If gamemode = registered And choice > 2 Then
        I_Printf "M_Episode: 4th episode requires UltimateDOOM"
        choice = 0
    End If

    epi = choice
    M_SetupNextMenu NewDef
End Sub

' M_Options
Private Sub M_DrawOptions()
    V_DrawPatchDirect 108, 15, 0, W_GetNumForName("M_OPTTTL")

    V_DrawPatchDirect menus(OptionsDef).x + 175, menus(OptionsDef).y + LINEHEIGHT * opt_detail, 0, _
                      W_GetNumForName(detailNames(detailLevel And 1))

    V_DrawPatchDirect menus(OptionsDef).x + 120, menus(OptionsDef).y + LINEHEIGHT * opt_messages, 0, _
                      W_GetNumForName(msgNames(showMessages And 1))

    M_DrawThermo menus(OptionsDef).x, menus(OptionsDef).y + LINEHEIGHT * (opt_mousesens + 1), 10, mouseSensitivity

    M_DrawThermo menus(OptionsDef).x, menus(OptionsDef).y + LINEHEIGHT * (opt_scrnsize + 1), 9, screenSize
End Sub

Private Sub M_Options(ByVal choice As Long)
    M_SetupNextMenu OptionsDef
End Sub

' Toggle messages on/off
Private Sub M_ChangeMessages(ByVal choice As Long)
    showMessages = 1 - showMessages

    If showMessages = 0 Then
        players(consoleplayer).message = MSGOFF
    Else
        players(consoleplayer).message = MSGON
    End If

    message_dontfuckwithme = True
End Sub

' M_EndGame
Private Sub M_EndGameResponse(ByVal ch As Long)
    If ch <> 121 Then Exit Sub                   ' 'y'

    menus(currentMenu).lastOn = itemOn
    M_ClearMenus
    D_StartTitle
End Sub

Private Sub M_EndGame(ByVal choice As Long)
    If Not usergame Then
        S_StartSound 0, sfx_oof
        Exit Sub
    End If

    If netgame Then
        M_StartMessage NETEND, MR_NONE, False
        Exit Sub
    End If

    M_StartMessage ENDGAME, MR_EndGameResponse, True
End Sub

' M_ReadThis
Private Sub M_ReadThis(ByVal choice As Long)
    M_SetupNextMenu ReadDef1
End Sub

Private Sub M_ReadThis2(ByVal choice As Long)
    M_SetupNextMenu ReadDef2
End Sub

Private Sub M_FinishReadThis(ByVal choice As Long)
    M_SetupNextMenu MainDef
End Sub

' M_QuitDOOM
Private Sub M_QuitResponse(ByVal ch As Long)
    If ch <> 121 Then Exit Sub                   ' 'y'
    If Not netgame Then
        If gamemode = commercial Then
            S_StartSound 0, quitsounds2((gametic \ 4) And 7)
        Else
            S_StartSound 0, quitsounds((gametic \ 4) And 7)
        End If
        I_WaitVBL 105
    End If
    I_Quit
End Sub

Private Sub M_QuitDOOM(ByVal choice As Long)
    Dim n As Long

    ' We pick index 0 which is language sensitive,
    '  or one at random, between 1 and maximum number.
    n = gametic Mod 8
    If language <> english Then
        n = 0
    ElseIf gamemode = commercial And n <> 0 Then
        n = n + 7                                ' the DOOM 2 messages
    End If
    endstring = endmsg(n) & vbLf & vbLf & DOSY

    M_StartMessage endstring, MR_QuitResponse, True
End Sub

Private Sub M_ChangeSensitivity(ByVal choice As Long)
    Select Case choice
        Case 0
            If mouseSensitivity <> 0 Then mouseSensitivity = mouseSensitivity - 1
        Case 1
            If mouseSensitivity < 9 Then mouseSensitivity = mouseSensitivity + 1
    End Select
End Sub

' Deviation: linuxdoom left low detail switched off here ("FIXME - does not
' work"); the renderer port has the low detail drawers, so this does what
' the DOS version did.
Private Sub M_ChangeDetail(ByVal choice As Long)
    detailLevel = 1 - detailLevel

    R_SetViewSize screenblocks, detailLevel

    If detailLevel = 0 Then
        players(consoleplayer).message = DETAILHI
    Else
        players(consoleplayer).message = DETAILLO
    End If
End Sub

Private Sub M_SizeDisplay(ByVal choice As Long)
    Select Case choice
        Case 0
            If screenSize > 0 Then
                screenblocks = screenblocks - 1
                screenSize = screenSize - 1
            End If
        Case 1
            If screenSize < 8 Then
                screenblocks = screenblocks + 1
                screenSize = screenSize + 1
            End If
    End Select

    R_SetViewSize screenblocks, detailLevel
End Sub

' Menu Functions
Private Sub M_DrawThermo(ByVal x As Long, ByVal y As Long, ByVal thermWidth As Long, ByVal thermDot As Long)
    Dim xx As Long, i As Long

    xx = x
    V_DrawPatchDirect xx, y, 0, W_GetNumForName("M_THERML")
    xx = xx + 8
    For i = 0 To thermWidth - 1
        V_DrawPatchDirect xx, y, 0, W_GetNumForName("M_THERMM")
        xx = xx + 8
    Next
    V_DrawPatchDirect xx, y, 0, W_GetNumForName("M_THERMR")

    V_DrawPatchDirect (x + 8) + thermDot * 8, y, 0, W_GetNumForName("M_THERMO")
End Sub

' (M_DrawEmptyCell and M_DrawSelCell are not used by the C code either.)

Private Sub M_StartMessage(ByVal s As String, ByVal routine As Long, ByVal needsInput As Boolean)
    messageLastMenuActive = menuactive
    messageToPrint = 1
    messageString = s
    messageRoutine = routine
    messageNeedsInput = needsInput
    menuactive = True
End Sub

Private Sub M_StopMessage()
    menuactive = messageLastMenuActive
    messageToPrint = 0
End Sub

' Find string width from hu_font chars
Private Function M_StringWidth(ByVal s As String) As Long
    Dim i As Long, w As Long, c As Long
    For i = 1 To Len(s)
        c = toupper(Asc(Mid$(s, i, 1))) - HU_FONTSTART
        If c < 0 Or c >= HU_FONTSIZE Then
            w = w + 4
        Else
            w = w + PatchWidth(hu_font(c))
        End If
    Next
    M_StringWidth = w
End Function

' Find string height from hu_font chars
Private Function M_StringHeight(ByVal s As String) As Long
    Dim i As Long, h As Long, height As Long
    height = PatchHeight(hu_font(0))
    h = height
    For i = 1 To Len(s)
        If Mid$(s, i, 1) = vbLf Then h = h + height
    Next
    M_StringHeight = h
End Function

' Write a string using the hu_font
Private Sub M_WriteText(ByVal x As Long, ByVal y As Long, ByVal s As String)
    Dim i As Long, w As Long, c As Long, cx As Long, cy As Long

    cx = x
    cy = y
    For i = 1 To Len(s)
        c = Asc(Mid$(s, i, 1))
        If c = 10 Then                           ' '\n'
            cx = x
            cy = cy + 12
        Else
            c = toupper(c) - HU_FONTSTART
            If c < 0 Or c >= HU_FONTSIZE Then
                cx = cx + 4
            Else
                w = PatchWidth(hu_font(c))
                If cx + w > SCREENWIDTH Then Exit For
                V_DrawPatchDirect cx, cy, 0, hu_font(c)
                cx = cx + w
            End If
        End If
    Next
End Sub

' CONTROL PANEL

' M_Responder
Public Function M_Responder(ev As event_t) As Boolean
    Dim ch As Long, i As Long, routine As Long, status As Long

    ch = -1

    If ev.typ = ev_joystick And joywait < I_GetTime() Then
        If ev.data3 = -1 Then
            ch = KEY_UPARROW
            joywait = I_GetTime() + 5
        ElseIf ev.data3 = 1 Then
            ch = KEY_DOWNARROW
            joywait = I_GetTime() + 5
        End If

        If ev.data2 = -1 Then
            ch = KEY_LEFTARROW
            joywait = I_GetTime() + 2
        ElseIf ev.data2 = 1 Then
            ch = KEY_RIGHTARROW
            joywait = I_GetTime() + 2
        End If

        If (ev.data1 And 1) <> 0 Then
            ch = KEY_ENTER
            joywait = I_GetTime() + 5
        End If
        If (ev.data1 And 2) <> 0 Then
            ch = KEY_BACKSPACE
            joywait = I_GetTime() + 5
        End If
    Else
        If ev.typ = ev_mouse And mousewait < I_GetTime() Then
            mousey = mousey + ev.data3
            If mousey < lasty - 30 Then
                ch = KEY_DOWNARROW
                mousewait = I_GetTime() + 5
                lasty = lasty - 30
                mousey = lasty
            ElseIf mousey > lasty + 30 Then
                ch = KEY_UPARROW
                mousewait = I_GetTime() + 5
                lasty = lasty + 30
                mousey = lasty
            End If

            mousex = mousex + ev.data2
            If mousex < lastx - 30 Then
                ch = KEY_LEFTARROW
                mousewait = I_GetTime() + 5
                lastx = lastx - 30
                mousex = lastx
            ElseIf mousex > lastx + 30 Then
                ch = KEY_RIGHTARROW
                mousewait = I_GetTime() + 5
                lastx = lastx + 30
                mousex = lastx
            End If

            If (ev.data1 And 1) <> 0 Then
                ch = KEY_ENTER
                mousewait = I_GetTime() + 15
            End If

            If (ev.data1 And 2) <> 0 Then
                ch = KEY_BACKSPACE
                mousewait = I_GetTime() + 15
            End If
        ElseIf ev.typ = ev_keydown Then
            ch = ev.data1
        End If
    End If

    If ch = -1 Then
        M_Responder = False
        Exit Function
    End If

    ' Save Game string input
    If saveStringEnter <> 0 Then
        Select Case ch
            Case KEY_BACKSPACE
                If saveCharIndex > 0 Then
                    saveCharIndex = saveCharIndex - 1
                    savegamestrings(saveSlot) = Left$(savegamestrings(saveSlot), saveCharIndex)
                End If

            Case KEY_ESCAPE
                saveStringEnter = 0
                savegamestrings(saveSlot) = saveOldString

            Case KEY_ENTER
                saveStringEnter = 0
                If savegamestrings(saveSlot) <> "" Then M_DoSave saveSlot

            Case Else
                ch = toupper(ch)
                If ch = 32 Or (ch - HU_FONTSTART >= 0 And ch - HU_FONTSTART < HU_FONTSIZE) Then
                    If ch >= 32 And ch <= 127 And saveCharIndex < SAVESTRINGSIZE - 1 Then
                        If M_StringWidth(savegamestrings(saveSlot)) < (SAVESTRINGSIZE - 2) * 8 Then
                            savegamestrings(saveSlot) = savegamestrings(saveSlot) & Chr$(ch)
                            saveCharIndex = saveCharIndex + 1
                        End If
                    End If
                End If
        End Select
        M_Responder = True
        Exit Function
    End If

    ' Take care of any messages that need input
    If messageToPrint <> 0 Then
        If messageNeedsInput And Not (ch = 32 Or ch = 110 Or ch = 121 Or ch = KEY_ESCAPE) Then
            M_Responder = False                  ' only ' ', 'n', 'y' and Esc answer
            Exit Function
        End If

        menuactive = messageLastMenuActive
        messageToPrint = 0
        If messageRoutine <> MR_NONE Then CallMessageRoutine messageRoutine, ch

        menuactive = False
        S_StartSound 0, sfx_swtchx
        M_Responder = True
        Exit Function
    End If

    If devparm And ch = KEY_F1 Then
        G_ScreenShot
        M_Responder = True
        Exit Function
    End If

    ' F-Keys
    If Not menuactive Then
        Select Case ch
            Case KEY_MINUS                       ' Screen size down
                If automapactive Or chat_on Then
                    M_Responder = False
                    Exit Function
                End If
                M_SizeDisplay 0
                S_StartSound 0, sfx_stnmov
                M_Responder = True
                Exit Function

            Case KEY_EQUALS                      ' Screen size up
                If automapactive Or chat_on Then
                    M_Responder = False
                    Exit Function
                End If
                M_SizeDisplay 1
                S_StartSound 0, sfx_stnmov
                M_Responder = True
                Exit Function

            Case KEY_F1                          ' Help key
                M_StartControlPanel

                If gamemode = retail Then
                    currentMenu = ReadDef2
                Else
                    currentMenu = ReadDef1
                End If

                itemOn = 0
                S_StartSound 0, sfx_swtchn
                M_Responder = True
                Exit Function

            Case KEY_F2                          ' Save
                M_StartControlPanel
                S_StartSound 0, sfx_swtchn
                M_SaveGame 0
                M_Responder = True
                Exit Function

            Case KEY_F3                          ' Load
                M_StartControlPanel
                S_StartSound 0, sfx_swtchn
                M_LoadGame 0
                M_Responder = True
                Exit Function

            Case KEY_F4                          ' Sound Volume
                M_StartControlPanel
                currentMenu = SoundDef
                itemOn = sfx_vol
                S_StartSound 0, sfx_swtchn
                M_Responder = True
                Exit Function

            Case KEY_F5                          ' Detail toggle
                M_ChangeDetail 0
                S_StartSound 0, sfx_swtchn
                M_Responder = True
                Exit Function

            Case KEY_F6                          ' Quicksave
                S_StartSound 0, sfx_swtchn
                M_QuickSave
                M_Responder = True
                Exit Function

            Case KEY_F7                          ' End game
                S_StartSound 0, sfx_swtchn
                M_EndGame 0
                M_Responder = True
                Exit Function

            Case KEY_F8                          ' Toggle messages
                M_ChangeMessages 0
                S_StartSound 0, sfx_swtchn
                M_Responder = True
                Exit Function

            Case KEY_F9                          ' Quickload
                S_StartSound 0, sfx_swtchn
                M_QuickLoad
                M_Responder = True
                Exit Function

            Case KEY_F10                         ' Quit DOOM
                S_StartSound 0, sfx_swtchn
                M_QuitDOOM 0
                M_Responder = True
                Exit Function

            Case KEY_F11                         ' gamma toggle
                usegamma = usegamma + 1
                If usegamma > 4 Then usegamma = 0
                players(consoleplayer).message = gammamsg(usegamma)
                I_SetPalette W_CacheLumpName("PLAYPAL")
                M_Responder = True
                Exit Function
        End Select
    End If

    ' Pop-up menu?
    If Not menuactive Then
        If ch = KEY_ESCAPE Then
            M_StartControlPanel
            S_StartSound 0, sfx_swtchn
            M_Responder = True
            Exit Function
        End If
        M_Responder = False
        Exit Function
    End If

    ' Keys usable within menu
    routine = menus(currentMenu).menuitems(itemOn).routine
    status = menus(currentMenu).menuitems(itemOn).status
    Select Case ch
        Case KEY_DOWNARROW
            Do
                If itemOn + 1 > menus(currentMenu).numitems - 1 Then
                    itemOn = 0
                Else
                    itemOn = itemOn + 1
                End If
                S_StartSound 0, sfx_pstop
            Loop While menus(currentMenu).menuitems(itemOn).status = -1
            M_Responder = True
            Exit Function

        Case KEY_UPARROW
            Do
                If itemOn = 0 Then
                    itemOn = menus(currentMenu).numitems - 1
                Else
                    itemOn = itemOn - 1
                End If
                S_StartSound 0, sfx_pstop
            Loop While menus(currentMenu).menuitems(itemOn).status = -1
            M_Responder = True
            Exit Function

        Case KEY_LEFTARROW
            If routine <> MI_NONE And status = 2 Then
                S_StartSound 0, sfx_stnmov
                CallItemRoutine routine, 0
            End If
            M_Responder = True
            Exit Function

        Case KEY_RIGHTARROW
            If routine <> MI_NONE And status = 2 Then
                S_StartSound 0, sfx_stnmov
                CallItemRoutine routine, 1
            End If
            M_Responder = True
            Exit Function

        Case KEY_ENTER
            If routine <> MI_NONE And status <> 0 Then
                menus(currentMenu).lastOn = itemOn
                If status = 2 Then
                    CallItemRoutine routine, 1   ' right arrow
                    S_StartSound 0, sfx_stnmov
                Else
                    CallItemRoutine routine, itemOn
                    S_StartSound 0, sfx_pistol
                End If
            End If
            M_Responder = True
            Exit Function

        Case KEY_ESCAPE
            menus(currentMenu).lastOn = itemOn
            M_ClearMenus
            S_StartSound 0, sfx_swtchx
            M_Responder = True
            Exit Function

        Case KEY_BACKSPACE
            menus(currentMenu).lastOn = itemOn
            If menus(currentMenu).prevMenu >= 0 Then
                currentMenu = menus(currentMenu).prevMenu
                itemOn = menus(currentMenu).lastOn
                S_StartSound 0, sfx_swtchn
            End If
            M_Responder = True
            Exit Function

        Case Else
            For i = itemOn + 1 To menus(currentMenu).numitems - 1
                If menus(currentMenu).menuitems(i).alphaKey = ch Then
                    itemOn = i
                    S_StartSound 0, sfx_pstop
                    M_Responder = True
                    Exit Function
                End If
            Next
            For i = 0 To itemOn
                If menus(currentMenu).menuitems(i).alphaKey = ch Then
                    itemOn = i
                    S_StartSound 0, sfx_pstop
                    M_Responder = True
                    Exit Function
                End If
            Next
    End Select

    M_Responder = False
End Function

' M_StartControlPanel
Public Sub M_StartControlPanel()
    ' intro might call this repeatedly
    If menuactive Then Exit Sub

    menuactive = True
    currentMenu = MainDef                        ' JDC
    itemOn = menus(currentMenu).lastOn           ' JDC
End Sub

' M_Drawer - called after the view has been rendered, but before it has
' been blitted.
Public Sub M_Drawer()
    Dim x As Long, y As Long, i As Long, start As Long, nl As Long, s As String

    inhelpscreens = False

    ' Horiz. & Vertically center string and print it.
    If messageToPrint <> 0 Then
        start = 1
        y = 100 - M_StringHeight(messageString) \ 2
        Do While start <= Len(messageString)
            nl = InStr(start, messageString, vbLf)
            If nl > 0 Then
                s = Mid$(messageString, start, nl - start)
                start = nl + 1
            Else
                s = Mid$(messageString, start)
                start = Len(messageString) + 1
            End If

            x = 160 - M_StringWidth(s) \ 2
            M_WriteText x, y, s
            y = y + PatchHeight(hu_font(0))
        Loop
        Exit Sub
    End If

    If Not menuactive Then Exit Sub

    If menus(currentMenu).routine <> MD_NONE Then CallDrawRoutine menus(currentMenu).routine ' call Draw routine

    ' DRAW MENU
    x = menus(currentMenu).x
    y = menus(currentMenu).y

    For i = 0 To menus(currentMenu).numitems - 1
        If menus(currentMenu).menuitems(i).name <> "" Then
            V_DrawPatchDirect x, y, 0, W_GetNumForName(menus(currentMenu).menuitems(i).name)
        End If
        y = y + LINEHEIGHT
    Next

    ' DRAW SKULL
    V_DrawPatchDirect x + SKULLXOFF, menus(currentMenu).y - 5 + itemOn * LINEHEIGHT, 0, _
                      W_GetNumForName(skullName(whichSkull))
End Sub

' M_ClearMenus
Private Sub M_ClearMenus()
    menuactive = False
    ' if (!netgame && usergame && paused)
    '       sendpause = true;
End Sub

' M_SetupNextMenu
Private Sub M_SetupNextMenu(ByVal menudef As Long)
    currentMenu = menudef
    itemOn = menus(currentMenu).lastOn
End Sub

' M_Ticker
Public Sub M_Ticker()
    skullAnimCounter = skullAnimCounter - 1
    If skullAnimCounter <= 0 Then
        whichSkull = whichSkull Xor 1
        skullAnimCounter = 8
    End If
End Sub

' M_Init
Public Sub M_Init()
    InitMenus

    currentMenu = MainDef
    menuactive = False
    itemOn = menus(currentMenu).lastOn
    whichSkull = 0
    skullAnimCounter = 10
    screenSize = screenblocks - 3
    messageToPrint = 0
    messageString = ""
    messageLastMenuActive = menuactive
    quickSaveSlot = -1

    ' Here we could catch other version dependencies,
    '  like HELP1/2, and four episodes.

    Select Case gamemode
        Case commercial
            ' This is used because DOOM 2 had only one HELP
            '  page. I use CREDIT as second page now, but
            '  kept this hack for educational purposes.
            menus(MainDef).menuitems(main_readthis) = menus(MainDef).menuitems(main_quitdoom)
            menus(MainDef).numitems = menus(MainDef).numitems - 1
            menus(MainDef).y = menus(MainDef).y + 8
            menus(NewDef).prevMenu = MainDef
            menus(ReadDef1).routine = MD_DrawReadThis1
            menus(ReadDef1).x = 330
            menus(ReadDef1).y = 165
            menus(ReadDef1).menuitems(0).routine = MI_FinishReadThis
        Case shareware, registered
            ' Episode 2 and 3 are handled,
            '  branching to an ad screen.
            ' We need to remove the fourth episode.
            menus(EpiDef).numitems = menus(EpiDef).numitems - 1
        Case Else
            ' retail: We are fine.
    End Select
End Sub

' dispatchers for the C function pointers

Private Sub CallItemRoutine(ByVal routine As Long, ByVal choice As Long)
    Select Case routine
        Case MI_NewGame: M_NewGame choice
        Case MI_Options: M_Options choice
        Case MI_LoadGame: M_LoadGame choice
        Case MI_SaveGame: M_SaveGame choice
        Case MI_ReadThis: M_ReadThis choice
        Case MI_QuitDOOM: M_QuitDOOM choice
        Case MI_Episode: M_Episode choice
        Case MI_ChooseSkill: M_ChooseSkill choice
        Case MI_EndGame: M_EndGame choice
        Case MI_ChangeMessages: M_ChangeMessages choice
        Case MI_ChangeDetail: M_ChangeDetail choice
        Case MI_SizeDisplay: M_SizeDisplay choice
        Case MI_ChangeSensitivity: M_ChangeSensitivity choice
        Case MI_Sound: M_Sound choice
        Case MI_ReadThis2: M_ReadThis2 choice
        Case MI_FinishReadThis: M_FinishReadThis choice
        Case MI_SfxVol: M_SfxVol choice
        Case MI_MusicVol: M_MusicVol choice
        Case MI_LoadSelect: M_LoadSelect choice
        Case MI_SaveSelect: M_SaveSelect choice
    End Select
End Sub

Private Sub CallDrawRoutine(ByVal routine As Long)
    Select Case routine
        Case MD_DrawMainMenu: M_DrawMainMenu
        Case MD_DrawEpisode: M_DrawEpisode
        Case MD_DrawNewGame: M_DrawNewGame
        Case MD_DrawOptions: M_DrawOptions
        Case MD_DrawReadThis1: M_DrawReadThis1
        Case MD_DrawReadThis2: M_DrawReadThis2
        Case MD_DrawSound: M_DrawSound
        Case MD_DrawLoad: M_DrawLoad
        Case MD_DrawSave: M_DrawSave
    End Select
End Sub

Private Sub CallMessageRoutine(ByVal routine As Long, ByVal ch As Long)
    Select Case routine
        Case MR_QuickSaveResponse: M_QuickSaveResponse ch
        Case MR_QuickLoadResponse: M_QuickLoadResponse ch
        Case MR_VerifyNightmare: M_VerifyNightmare ch
        Case MR_EndGameResponse: M_EndGameResponse ch
        Case MR_QuitResponse: M_QuitResponse ch
    End Select
End Sub
