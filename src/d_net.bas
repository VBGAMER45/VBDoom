Attribute VB_Name = "d_net"
' d_net.c - DOOM network game communication and protocol,
'  all OS independent parts.
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
'
' Lockstep, peer to peer: every node sends its ticcmds to every other node,
' and a gametic runs only when the commands of all players have arrived.
' Node 0 is this computer; packets to it "rebound" without touching the
' network, so a single player game takes the same path.
' (C: doomcom_t* doomcom and netbuffer = &doomcom->data; here one doomcom
' record whose data member is the netbuffer.)
Option Explicit

Public Const DOOMCOM_ID As Long = &H12345678
Public Const MAXNETNODES As Long = 8
Public Const BACKUPTICS As Long = 12

' command_t
Public Const CMD_SEND As Long = 1
Public Const CMD_GET As Long = 2

Public Type doomdata_t
    ' High bit is retransmit request.
    checksum As Long
    ' Only valid if NCMD_RETRANSMIT.
    retransmitfrom As Long                       ' byte
    starttic As Long                             ' byte
    player As Long                               ' byte
    numtics As Long                              ' byte
    cmds(0 To BACKUPTICS - 1) As ticcmd_t
End Type

' (the C's intnum, deathmatch, savegame, episode, map, skill, angleoffset
' and drone fields were not used by linuxdoom)
Public Type doomcom_t
    ' Supposed to be DOOMCOM_ID?
    id As Long
    ' Communication between DOOM and the driver.
    ' Is CMD_SEND or CMD_GET.
    command As Long
    ' Is dest for send, set by get (-1 = no packet).
    remotenode As Long
    ' Number of bytes in doomdata to be sent
    datalength As Long
    ' Info common to all nodes.
    ' Console is allways node 0.
    numnodes As Long
    ' Flag: 1 = no duplication, 2-5 = dup for slow nets.
    ticdup As Long
    ' Flag: 1 = send a backup tic in every packet.
    extratics As Long
    ' Info specific to this node.
    consoleplayer As Long
    numplayers As Long
    ' The packet data to be sent.
    data As doomdata_t
End Type

Private Const NCMD_EXIT As Long = &H80000000
Private Const NCMD_RETRANSMIT As Long = &H40000000
Private Const NCMD_SETUP As Long = &H20000000
Private Const NCMD_KILL As Long = &H10000000     ' kill game
Private Const NCMD_CHECKSUM As Long = &HFFFFFFF

Private Const RESEND_COUNT As Long = 10        ' C: RESENDCOUNT
Private Const PL_DRONE As Long = &H80            ' bit flag in doomdata->player

Private Const MAXINT As Long = &H7FFFFFFF

Public doomcom As doomcom_t

'
' NETWORKING
'
' gametic is the tic about to (or currently being) run
' maketic is the tick that hasn't had control made for it yet
' nettics[] has the maketics for all players
'
' a gametic cannot be run until nettics[] > gametic for all players
'
Private localcmds(0 To BACKUPTICS - 1) As ticcmd_t
Public netcmds(0 To MAXPLAYERS - 1, 0 To BACKUPTICS - 1) As ticcmd_t
Private nettics(0 To MAXNETNODES - 1) As Long
Private nodeingame(0 To MAXNETNODES - 1) As Boolean    ' set false as nodes leave game
Private remoteresend(0 To MAXNETNODES - 1) As Boolean  ' set when local needs tics
Private resendto(0 To MAXNETNODES - 1) As Long         ' set when remote needs tics
Private resendcount(0 To MAXNETNODES - 1) As Long

Private nodeforplayer(0 To MAXPLAYERS - 1) As Long

Public maketic As Long
Private skiptics As Long
Public ticdup As Long
Private maxsend As Long                          ' BACKUPTICS/(2*ticdup)-1

Private reboundpacket As Boolean
Private reboundstore As doomdata_t

Public debugfile As Integer                      ' -debugfile (0 = none)

Private gametime As Long

Private frameon As Long
Private frameskip(0 To 3) As Boolean
Private oldnettics As Long

'
'
'
Private Function NetbufferSize() As Long
    ' the 8 byte header, then 8 bytes per ticcmd
    NetbufferSize = 8 + 8 * doomcom.data.numtics
End Function

'
' Checksum
' (linuxdoom returned 0: byte order problems.  UDP checks its own data.)
'
Private Function NetbufferChecksum() As Long
    NetbufferChecksum = 0
End Function

'
'
'
Private Function ExpandTics(ByVal low As Long) As Long
    Dim delta As Long

    delta = low - (maketic And &HFF&)

    If delta >= -64 And delta <= 64 Then
        ExpandTics = (maketic And Not &HFF&) + low
    ElseIf delta > 64 Then
        ExpandTics = (maketic And Not &HFF&) - 256 + low
    Else
        ExpandTics = (maketic And Not &HFF&) + 256 + low
    End If
End Function

' C's % for a tic number that can be -1 (resendto with -extratic)
Private Function TicSlot(ByVal tic As Long) As Long
    TicSlot = ((tic Mod BACKUPTICS) + BACKUPTICS) Mod BACKUPTICS
End Function

' The C fields of ticcmd_t are char, char, short, short, byte, byte: wrap the
' values the same way, so the local copy equals what the other nodes get.
Private Sub NormalizeTiccmd(cmd As ticcmd_t)
    cmd.forwardmove = cmd.forwardmove And &HFF&
    If cmd.forwardmove >= 128 Then cmd.forwardmove = cmd.forwardmove - 256
    cmd.sidemove = cmd.sidemove And &HFF&
    If cmd.sidemove >= 128 Then cmd.sidemove = cmd.sidemove - 256
    cmd.angleturn = cmd.angleturn And &HFFFF&
    If cmd.angleturn >= 32768 Then cmd.angleturn = cmd.angleturn - 65536
    cmd.consistancy = cmd.consistancy And &HFFFF&
    If cmd.consistancy >= 32768 Then cmd.consistancy = cmd.consistancy - 65536
    cmd.chatchar = cmd.chatchar And &HFF&
    cmd.buttons = cmd.buttons And &HFF&
End Sub

Private Sub DebugPacket(ByVal what As String)
    Dim realretrans As Long

    If (doomcom.data.checksum And NCMD_RETRANSMIT) <> 0 Then
        realretrans = ExpandTics(doomcom.data.retransmitfrom)
    Else
        realretrans = -1
    End If
    Print #debugfile, what & " (" & ExpandTics(doomcom.data.starttic) & " + " & _
        doomcom.data.numtics & ", R " & realretrans & ") [" & doomcom.datalength & "]"
End Sub

'
' HSendPacket
'
Private Sub HSendPacket(ByVal node As Long, ByVal flags As Long)
    doomcom.data.checksum = NetbufferChecksum() Or flags

    If node = 0 Then
        reboundstore = doomcom.data
        reboundpacket = True
        Exit Sub
    End If

    If demoplayback Then Exit Sub

    If Not netgame Then I_Error "Tried to transmit to another node"

    doomcom.command = CMD_SEND
    doomcom.remotenode = node
    doomcom.datalength = NetbufferSize()

    If debugfile <> 0 Then DebugPacket "send"

    I_NetCmd
End Sub

'
' HGetPacket
' Returns false if no packet is waiting
'
Private Function HGetPacket() As Boolean
    If reboundpacket Then
        doomcom.data = reboundstore
        doomcom.remotenode = 0
        reboundpacket = False
        HGetPacket = True
        Exit Function
    End If

    If Not netgame Then Exit Function

    If demoplayback Then Exit Function

    doomcom.command = CMD_GET
    I_NetCmd

    If doomcom.remotenode = -1 Then Exit Function

    If doomcom.datalength <> NetbufferSize() Then
        If debugfile <> 0 Then Print #debugfile, "bad packet length " & doomcom.datalength
        Exit Function
    End If

    If NetbufferChecksum() <> (doomcom.data.checksum And NCMD_CHECKSUM) Then
        If debugfile <> 0 Then Print #debugfile, "bad packet checksum"
        Exit Function
    End If

    If debugfile <> 0 Then
        If (doomcom.data.checksum And NCMD_SETUP) <> 0 Then
            Print #debugfile, "setup packet"
        Else
            DebugPacket "get " & doomcom.remotenode & " ="
        End If
    End If

    HGetPacket = True
End Function

'
' GetPackets
'
Private Sub GetPackets()
    Dim netconsole As Long, netnode As Long, realend As Long, realstart As Long
    Dim src As Long

    Do While HGetPacket()
        If (doomcom.data.checksum And NCMD_SETUP) <> 0 Then GoTo nextpacket   ' extra setup packet

        netconsole = doomcom.data.player And Not PL_DRONE
        netnode = doomcom.remotenode

        ' (the C indexed its arrays with whatever the packet said)
        If netconsole >= MAXPLAYERS Then GoTo nextpacket

        ' to save bytes, only the low byte of tic numbers are sent
        ' Figure out what the rest of the bytes are
        realstart = ExpandTics(doomcom.data.starttic)
        realend = realstart + doomcom.data.numtics

        ' check for exiting the game
        If (doomcom.data.checksum And NCMD_EXIT) <> 0 Then
            If Not nodeingame(netnode) Then GoTo nextpacket
            nodeingame(netnode) = False
            playeringame(netconsole) = False
            players(consoleplayer).message = "Player " & (netconsole + 1) & " left the game"
            If demorecording Then G_CheckDemoStatus
            GoTo nextpacket
        End If

        ' check for a remote game kill
        If (doomcom.data.checksum And NCMD_KILL) <> 0 Then I_Error "Killed by network driver"

        nodeforplayer(netconsole) = netnode

        ' check for retransmit request
        If resendcount(netnode) <= 0 And (doomcom.data.checksum And NCMD_RETRANSMIT) <> 0 Then
            resendto(netnode) = ExpandTics(doomcom.data.retransmitfrom)
            If debugfile <> 0 Then Print #debugfile, "retransmit from " & resendto(netnode)
            resendcount(netnode) = RESEND_COUNT
        Else
            resendcount(netnode) = resendcount(netnode) - 1
        End If

        ' check for out of order / duplicated packet
        If realend = nettics(netnode) Then GoTo nextpacket

        If realend < nettics(netnode) Then
            If debugfile <> 0 Then
                Print #debugfile, "out of order packet (" & realstart & " + " & doomcom.data.numtics & ")"
            End If
            GoTo nextpacket
        End If

        ' check for a missed packet
        If realstart > nettics(netnode) Then
            ' stop processing until the other system resends the missed tics
            If debugfile <> 0 Then
                Print #debugfile, "missed tics from " & netnode & " (" & realstart & " - " & nettics(netnode) & ")"
            End If
            remoteresend(netnode) = True
            GoTo nextpacket
        End If

        ' update command store from the packet
        remoteresend(netnode) = False

        src = nettics(netnode) - realstart

        Do While nettics(netnode) < realend
            netcmds(netconsole, nettics(netnode) Mod BACKUPTICS) = doomcom.data.cmds(src)
            nettics(netnode) = nettics(netnode) + 1
            src = src + 1
        Loop
nextpacket:
    Loop
End Sub

'
' NetUpdate
' Builds ticcmds for console player,
' sends out a packet
'
Public Sub NetUpdate()
    Dim nowtime As Long, newtics As Long, i As Long, j As Long
    Dim realstart As Long, gameticdiv As Long

    ' check time
    nowtime = I_GetTime() \ ticdup
    newtics = nowtime - gametime
    gametime = nowtime

    If newtics <= 0 Then GoTo listen             ' nothing new to update

    If skiptics <= newtics Then
        newtics = newtics - skiptics
        skiptics = 0
    Else
        skiptics = skiptics - newtics
        newtics = 0
    End If

    doomcom.data.player = consoleplayer

    ' build new ticcmds for console player
    gameticdiv = gametic \ ticdup
    For i = 0 To newtics - 1
        I_StartTic
        D_ProcessEvents
        If maketic - gameticdiv >= BACKUPTICS \ 2 - 1 Then Exit For   ' can't hold any more

        G_BuildTiccmd localcmds(maketic Mod BACKUPTICS)
        NormalizeTiccmd localcmds(maketic Mod BACKUPTICS)
        maketic = maketic + 1
    Next

    If singletics Then Exit Sub                  ' singletic update is syncronous

    ' send the packet to the other nodes
    For i = 0 To doomcom.numnodes - 1
        If nodeingame(i) Then
            realstart = resendto(i)
            doomcom.data.starttic = realstart And &HFF&
            doomcom.data.numtics = maketic - realstart
            If doomcom.data.numtics > BACKUPTICS Then I_Error "NetUpdate: netbuffer->numtics > BACKUPTICS"

            resendto(i) = maketic - doomcom.extratics

            For j = 0 To doomcom.data.numtics - 1
                doomcom.data.cmds(j) = localcmds(TicSlot(realstart + j))
            Next

            If remoteresend(i) Then
                doomcom.data.retransmitfrom = nettics(i) And &HFF&
                HSendPacket i, NCMD_RETRANSMIT
            Else
                doomcom.data.retransmitfrom = 0
                HSendPacket i, 0
            End If
        End If
    Next

listen:
    ' listen for other packets
    GetPackets
End Sub

'
' CheckAbort
'
Private Sub CheckAbort()
    Dim stoptic As Long

    stoptic = I_GetTime() + 2
    Do While I_GetTime() < stoptic
        I_StartTic
        Sleep 1                                  ' (Win32: do not spin)
    Loop

    I_StartTic
    Do While eventtail <> eventhead
        If events(eventtail).typ = ev_keydown And events(eventtail).data1 = KEY_ESCAPE Then
            I_Error "Network game synchronization aborted."
        End If
        eventtail = (eventtail + 1) And (MAXEVENTS - 1)
    Loop
End Sub

'
' D_ArbitrateNetStart
'
Private Sub D_ArbitrateNetStart()
    Dim i As Long
    Dim gotinfo(0 To MAXNETNODES - 1) As Boolean

    autostart = True

    ' (Win32: there is no console, so the window comes up early and its
    ' title says what the game waits for; Esc aborts)
    I_InitGraphics
    I_SetWindowTitle "DoomVB - waiting for the other players (Esc aborts)"

    If doomcom.consoleplayer <> 0 Then
        ' listen for setup info from key player
        I_Printf "listening for network start info..."
        Do
            CheckAbort
            If HGetPacket() Then
                If (doomcom.data.checksum And NCMD_SETUP) <> 0 Then
                    If doomcom.data.player <> VERSION Then
                        I_Error "Different DOOM versions cannot play a net game!"
                    End If
                    startskill = doomcom.data.retransmitfrom And 15
                    deathmatch = (doomcom.data.retransmitfrom And &HC0) \ 64
                    nomonsters = (doomcom.data.retransmitfrom And &H20) > 0
                    respawnparm = (doomcom.data.retransmitfrom And &H10) > 0
                    startmap = doomcom.data.starttic And &H3F
                    startepisode = doomcom.data.starttic \ 64
                    Exit Do
                End If
            End If
        Loop
    Else
        ' key player, send the setup info
        I_Printf "sending network start info..."
        Do
            CheckAbort
            For i = 0 To doomcom.numnodes - 1
                doomcom.data.retransmitfrom = startskill
                If deathmatch <> 0 Then
                    doomcom.data.retransmitfrom = doomcom.data.retransmitfrom Or (deathmatch * 64)
                End If
                If nomonsters Then doomcom.data.retransmitfrom = doomcom.data.retransmitfrom Or &H20
                If respawnparm Then doomcom.data.retransmitfrom = doomcom.data.retransmitfrom Or &H10
                doomcom.data.starttic = startepisode * 64 + startmap
                doomcom.data.player = VERSION
                doomcom.data.numtics = 0
                HSendPacket i, NCMD_SETUP
            Next

            i = 10
            Do While i > 0
                If Not HGetPacket() Then Exit Do
                If (doomcom.data.player And &H7F) < MAXNETNODES Then
                    gotinfo(doomcom.data.player And &H7F) = True
                End If
                i = i - 1
            Loop

            For i = 1 To doomcom.numnodes - 1
                If Not gotinfo(i) Then Exit For
            Next
        Loop While i < doomcom.numnodes
    End If

    I_SetWindowTitle ""
End Sub

'
' D_CheckNetGame
' Works out player numbers among the net participants
'
Public Sub D_CheckNetGame()
    Dim i As Long

    For i = 0 To MAXNETNODES - 1
        nodeingame(i) = False
        nettics(i) = 0
        remoteresend(i) = False                  ' set when local needs tics
        resendto(i) = 0                          ' which tic to start sending
    Next
    For i = 0 To MAXPLAYERS - 1
        playeringame(i) = False
    Next

    ' I_InitNetwork sets doomcom and netgame
    I_InitNetwork
    If doomcom.id <> DOOMCOM_ID Then I_Error "Doomcom buffer invalid!"

    consoleplayer = doomcom.consoleplayer
    displayplayer = doomcom.consoleplayer
    If netgame Then D_ArbitrateNetStart

    I_Printf "startskill " & startskill & "  deathmatch: " & deathmatch & _
        "  startmap: " & startmap & "  startepisode: " & startepisode

    ' read values out of doomcom
    ticdup = doomcom.ticdup
    maxsend = BACKUPTICS \ (2 * ticdup) - 1
    If maxsend < 1 Then maxsend = 1

    For i = 0 To doomcom.numplayers - 1
        playeringame(i) = True
    Next
    For i = 0 To doomcom.numnodes - 1
        nodeingame(i) = True
    Next

    I_Printf "player " & (consoleplayer + 1) & " of " & doomcom.numplayers & _
        " (" & doomcom.numnodes & " nodes)"

    ' (the C started from gametime 0: the first tics would pile up)
    gametime = I_GetTime() \ ticdup
End Sub

'
' D_QuitNetGame
' Called before quitting to leave a net game
' without hanging the other players
'
Public Sub D_QuitNetGame()
    Dim i As Long, j As Long

    If debugfile <> 0 Then
        Close #debugfile
        debugfile = 0
    End If

    If Not netgame Or Not usergame Or consoleplayer = -1 Or demoplayback Then Exit Sub

    ' send a bunch of packets for security
    doomcom.data.player = consoleplayer
    doomcom.data.numtics = 0
    For i = 0 To 3
        For j = 1 To doomcom.numnodes - 1
            If nodeingame(j) Then HSendPacket j, NCMD_EXIT
        Next
        I_WaitVBL 1
    Next
End Sub

'
' TryRunTics
'
Public Sub TryRunTics()
    Dim i As Long, j As Long, lowtic As Long, entertic As Long, realtics As Long
    Dim availabletics As Long, counts As Long, numplaying As Long, buf As Long
    Static oldentertics As Long

    ' get real tics
    entertic = I_GetTime() \ ticdup
    realtics = entertic - oldentertics
    oldentertics = entertic

    ' get available tics
    NetUpdate

    lowtic = MAXINT
    numplaying = 0
    For i = 0 To doomcom.numnodes - 1
        If nodeingame(i) Then
            numplaying = numplaying + 1
            If nettics(i) < lowtic Then lowtic = nettics(i)
        End If
    Next
    availabletics = lowtic - gametic \ ticdup

    ' decide how many tics to run
    If realtics < availabletics - 1 Then
        counts = realtics + 1
    ElseIf realtics < availabletics Then
        counts = realtics
    Else
        counts = availabletics
    End If

    If counts < 1 Then counts = 1

    frameon = frameon + 1

    If debugfile <> 0 Then
        Print #debugfile, "=======real: " & realtics & "  avail: " & availabletics & "  game: " & counts
    End If

    If Not demoplayback Then
        ' ideally nettics[0] should be 1 - 3 tics above lowtic
        ' if we are consistantly slower, speed up time
        For i = 0 To MAXPLAYERS - 1
            If playeringame(i) Then Exit For
        Next
        If consoleplayer = i Then
            ' the key player does not adapt
        Else
            If nettics(0) <= nettics(nodeforplayer(i)) Then
                gametime = gametime - 1
            End If
            frameskip(frameon And 3) = (oldnettics > nettics(nodeforplayer(i)))
            oldnettics = nettics(0)
            If frameskip(0) And frameskip(1) And frameskip(2) And frameskip(3) Then
                skiptics = 1
            End If
        End If
    End If

    ' wait for new tics if needed
    Do While lowtic < gametic \ ticdup + counts
        Sleep 1                                  ' (Win32: do not spin, and
        I_UpdateSound                            ' feed the sound: there is
        I_UpdateMusic                            ' no sound interrupt)
        NetUpdate
        lowtic = MAXINT

        For i = 0 To doomcom.numnodes - 1
            If nodeingame(i) And nettics(i) < lowtic Then lowtic = nettics(i)
        Next

        If lowtic < gametic \ ticdup Then I_Error "TryRunTics: lowtic < gametic"

        ' don't stay in here forever -- give the menu a chance to work
        If I_GetTime() \ ticdup - entertic >= 20 Then
            M_Ticker
            Exit Sub
        End If
    Loop

    ' run the count * ticdup dics
    Do While counts > 0
        counts = counts - 1
        For i = 0 To ticdup - 1
            If gametic \ ticdup > lowtic Then I_Error "gametic>lowtic"
            If advancedemo Then D_DoAdvanceDemo
            dbgPhase = "M_Ticker"
            M_Ticker
            dbgPhase = "G_Ticker"
            G_Ticker
            gametic = gametic + 1

            ' modify command for duplicated tics
            If i <> ticdup - 1 Then
                buf = (gametic \ ticdup) Mod BACKUPTICS
                For j = 0 To MAXPLAYERS - 1
                    netcmds(j, buf).chatchar = 0
                    If (netcmds(j, buf).buttons And BT_SPECIAL) <> 0 Then netcmds(j, buf).buttons = 0
                Next
            End If
        Next
        NetUpdate                                ' check for new console commands
    Loop
End Sub
