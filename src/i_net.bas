Attribute VB_Name = "i_net"
' i_net.c - UDP network driver, rewritten for Winsock
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
'
' As in linuxdoom:  -net <playernumber> <host> <host> ...  starts a net game
' with one host for each other player (a name, or ".1.2.3.4"), all on UDP
' port 5029 (-port n changes it).  The packet layout is linuxdoom's (the
' checksum, angleturn and consistancy in network byte order).
' Win32 changes:
'  * a host may give its port (host:port), and packets are sent from the
'    listening socket (linuxdoom used a second one), so a node is known by
'    address and port: several players can run on one computer, each with
'    its own -port.
'  * when every host is 127.x.x.x the socket only listens on 127.0.0.1, so a
'    local test opens no port on the network.
'  * a "connection reset" from a host that is not listening yet is ignored.
'  * test aid: -netdrop n drops every n-th packet sent.
Option Explicit

Private Type sockaddr_in
    sin_family As Integer
    sin_port As Integer                          ' network byte order
    sin_addr As Long                             ' network byte order
    sin_zero(0 To 7) As Byte
End Type

Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, lpWSAData As Any) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal stype As Long, ByVal protocol As Long) As Long
Private Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, addr As sockaddr_in, ByVal addrlen As Long) As Long
Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Private Declare Function ioctlsocket Lib "ws2_32.dll" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare Function sendto Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, _
    ByVal flags As Long, toaddr As sockaddr_in, ByVal tolen As Long) As Long
Private Declare Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, _
    ByVal flags As Long, fromaddr As sockaddr_in, fromlen As Long) As Long
Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Private Declare Function gethostbyname Lib "ws2_32.dll" (ByVal hostname As String) As Long

Private Const AF_INET As Long = 2
Private Const SOCK_DGRAM As Long = 2
Private Const IPPROTO_UDP As Long = 17
Private Const INVALID_SOCKET As Long = -1
Private Const SOCKET_ERROR As Long = -1
Private Const INADDR_NONE As Long = -1
Private Const FIONBIO As Long = &H8004667E
Private Const WSAEWOULDBLOCK As Long = 10035
Private Const WSAEMSGSIZE As Long = 10040
Private Const WSAECONNRESET As Long = 10054
' (error codes come from Err.LastDllError: the VB runtime overwrites the
' thread's last error before WSAGetLastError could read it)

'
' NETWORKING
'
Private DOOMPORT As Long                         ' (IPPORT_USERRESERVED + 0x1d)

Private insocket As Long                         ' also sends (see above)
Private sendaddress(0 To MAXNETNODES - 1) As sockaddr_in
Private wsaup As Boolean

Private pkt(0 To 8 + 8 * BACKUPTICS - 1) As Byte
Private rbuf(0 To 1023) As Byte

Private netdrop As Long
Private sendcount As Long

' A port number in network byte order.
Private Function PortN(ByVal port As Long) As Integer
    Dim v As Long
    v = (port And &HFF&) * 256& Or ((port \ 256) And &HFF&)
    If v >= 32768 Then v = v - 65536
    PortN = v
End Function

Private Function ShortFrom(ByVal hi As Long, ByVal lo As Long) As Long
    ShortFrom = hi * 256& + lo
    If ShortFrom >= 32768 Then ShortFrom = ShortFrom - 65536
End Function

'
' UDPsocket
'
Private Function UDPsocket() As Long
    Dim s As Long

    ' allocate a socket
    s = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP)
    If s = INVALID_SOCKET Then I_Error "can't create socket: error " & Err.LastDllError
    UDPsocket = s
End Function

'
' BindToLocalPort
'
Private Sub BindToLocalPort(ByVal s As Long, ByVal port As Long, ByVal localonly As Boolean)
    Dim address As sockaddr_in

    address.sin_family = AF_INET
    If localonly Then
        address.sin_addr = inet_addr("127.0.0.1")
    Else
        address.sin_addr = 0                     ' INADDR_ANY
    End If
    address.sin_port = PortN(port)

    If bind(s, address, LenB(address)) = SOCKET_ERROR Then
        I_Error "BindToPort: bind: error " & Err.LastDllError & " (is UDP port " & port & " in use?)"
    End If
End Sub

'
' PacketSend
'
Private Sub PacketSend()
    Dim c As Long, o As Long, v As Long

    sendcount = sendcount + 1
    If netdrop > 0 Then
        If sendcount Mod netdrop = 0 Then Exit Sub
    End If

    ' byte swap
    v = doomcom.data.checksum
    pkt(0) = ((v And &HFF000000) \ &H1000000) And &HFF&
    pkt(1) = (v And &HFF0000) \ &H10000
    pkt(2) = (v And &HFF00&) \ &H100&
    pkt(3) = v And &HFF&
    pkt(4) = doomcom.data.retransmitfrom And &HFF&
    pkt(5) = doomcom.data.starttic And &HFF&
    pkt(6) = doomcom.data.player And &HFF&
    pkt(7) = doomcom.data.numtics And &HFF&

    o = 8
    For c = 0 To doomcom.data.numtics - 1
        pkt(o) = doomcom.data.cmds(c).forwardmove And &HFF&
        pkt(o + 1) = doomcom.data.cmds(c).sidemove And &HFF&
        pkt(o + 2) = (doomcom.data.cmds(c).angleturn And &HFF00&) \ &H100&
        pkt(o + 3) = doomcom.data.cmds(c).angleturn And &HFF&
        pkt(o + 4) = (doomcom.data.cmds(c).consistancy And &HFF00&) \ &H100&
        pkt(o + 5) = doomcom.data.cmds(c).consistancy And &HFF&
        pkt(o + 6) = doomcom.data.cmds(c).chatchar And &HFF&
        pkt(o + 7) = doomcom.data.cmds(c).buttons And &HFF&
        o = o + 8
    Next

    ' (errors are ignored, as in linuxdoom: a lost packet is sent again)
    sendto insocket, pkt(0), doomcom.datalength, 0, sendaddress(doomcom.remotenode), LenB(sendaddress(0))
End Sub

'
' PacketGet
'
Private Sub PacketGet()
    Dim i As Long, c As Long, n As Long, o As Long, e As Long
    Dim fromaddress As sockaddr_in, fromlen As Long

    Do
        fromlen = LenB(fromaddress)
        c = recvfrom(insocket, rbuf(0), UBound(rbuf) + 1, 0, fromaddress, fromlen)
        If c <> SOCKET_ERROR Then Exit Do

        e = Err.LastDllError
        If e = WSAEWOULDBLOCK Then
            doomcom.remotenode = -1              ' no packet
            Exit Sub
        End If
        ' a send to a host that is not listening (yet) comes back as a
        ' "connection reset" on the next receive; oversized datagrams are
        ' not DOOM packets
        If e <> WSAECONNRESET And e <> WSAEMSGSIZE Then I_Error "GetPacket: error " & e
    Loop

    ' find remote node number
    For i = 0 To doomcom.numnodes - 1
        If fromaddress.sin_addr = sendaddress(i).sin_addr Then
            If fromaddress.sin_port = sendaddress(i).sin_port Then Exit For
        End If
    Next

    If i = doomcom.numnodes Then
        ' packet is not from one of the players (new game broadcast)
        doomcom.remotenode = -1                  ' no packet
        Exit Sub
    End If

    ' (the C copied as many ticcmds as the packet claimed)
    n = 0
    If c >= 8 Then n = rbuf(7)
    If c < 8 Or n > BACKUPTICS Then
        doomcom.remotenode = -1
        Exit Sub
    End If

    doomcom.remotenode = i                       ' good packet from a game player
    doomcom.datalength = c

    ' byte swap
    doomcom.data.checksum = (rbuf(0) And &H7F) * &H1000000 + rbuf(1) * &H10000 + rbuf(2) * &H100& + rbuf(3)
    If rbuf(0) >= 128 Then doomcom.data.checksum = doomcom.data.checksum Or &H80000000
    doomcom.data.retransmitfrom = rbuf(4)
    doomcom.data.starttic = rbuf(5)
    doomcom.data.player = rbuf(6)
    doomcom.data.numtics = n

    o = 8
    For c = 0 To n - 1
        doomcom.data.cmds(c).forwardmove = rbuf(o)
        If rbuf(o) >= 128 Then doomcom.data.cmds(c).forwardmove = doomcom.data.cmds(c).forwardmove - 256
        doomcom.data.cmds(c).sidemove = rbuf(o + 1)
        If rbuf(o + 1) >= 128 Then doomcom.data.cmds(c).sidemove = doomcom.data.cmds(c).sidemove - 256
        doomcom.data.cmds(c).angleturn = ShortFrom(rbuf(o + 2), rbuf(o + 3))
        doomcom.data.cmds(c).consistancy = ShortFrom(rbuf(o + 4), rbuf(o + 5))
        doomcom.data.cmds(c).chatchar = rbuf(o + 6)
        doomcom.data.cmds(c).buttons = rbuf(o + 7)
        o = o + 8
    Next
End Sub

' A host name or dotted address (network byte order).
Private Function HostAddress(ByVal host As String) As Long
    Dim a As Long, he As Long, lst As Long, first As Long

    a = inet_addr(host)
    If a = INADDR_NONE Then
        he = gethostbyname(host)
        If he = 0 Then I_Error "gethostbyname: couldn't find " & host
        ' hostent: h_name, h_aliases, h_addrtype + h_length, h_addr_list
        CopyMemory lst, ByVal he + 12, 4
        CopyMemory first, ByVal lst, 4
        CopyMemory a, ByVal first, 4
    End If
    HostAddress = a
End Function

'
' I_InitNetwork
'
Public Sub I_InitNetwork()
    Dim i As Long, p As Long, k As Long, port As Long, nb As Long
    Dim host As String, localonly As Boolean
    Dim wsadata(0 To 511) As Byte

    doomcom.id = 0
    doomcom.numnodes = 0

    ' set up for network
    i = M_CheckParm("-dup")
    If i > 0 And i < myargc - 1 Then
        doomcom.ticdup = Asc(myargv(i + 1)) - Asc("0")
        If doomcom.ticdup < 1 Then doomcom.ticdup = 1
        If doomcom.ticdup > 9 Then doomcom.ticdup = 9
    Else
        doomcom.ticdup = 1
    End If

    If M_CheckParm("-extratic") > 0 Then
        doomcom.extratics = 1
    Else
        doomcom.extratics = 0
    End If

    DOOMPORT = 5000 + &H1D
    p = M_CheckParm("-port")
    If p > 0 And p < myargc - 1 Then
        DOOMPORT = Val(myargv(p + 1))
        I_Printf "using alternate port " & DOOMPORT
    End If

    p = M_CheckParm("-netdrop")
    If p > 0 And p < myargc - 1 Then netdrop = Val(myargv(p + 1))

    ' parse network game options,
    '  -net <consoleplayer> <host> <host> ...
    i = M_CheckParm("-net")
    If i = 0 Then
        ' single player game
        netgame = False
        doomcom.id = DOOMCOM_ID
        doomcom.numplayers = 1
        doomcom.numnodes = 1
        doomcom.consoleplayer = 0
        Exit Sub
    End If

    netgame = True

    If WSAStartup(&H202, wsadata(0)) <> 0 Then I_Error "I_InitNetwork: no Winsock"
    wsaup = True

    ' parse player number and host list
    If i >= myargc - 1 Then I_Error "-net: player number missing"
    doomcom.consoleplayer = Asc(myargv(i + 1)) - Asc("1")
    doomcom.numnodes = 1                         ' this node for sure
    localonly = True
    i = i + 1
    Do
        i = i + 1
        If i >= myargc Then Exit Do
        If Left$(myargv(i), 1) = "-" Then Exit Do
        If doomcom.numnodes >= MAXPLAYERS Then I_Error "-net: at most " & MAXPLAYERS & " players"

        host = myargv(i)
        port = DOOMPORT
        k = InStr(host, ":")
        If k > 0 Then
            port = Val(Mid$(host, k + 1))
            host = Left$(host, k - 1)
        End If
        If Left$(host, 1) = "." Then host = Mid$(host, 2)

        sendaddress(doomcom.numnodes).sin_family = AF_INET
        sendaddress(doomcom.numnodes).sin_port = PortN(port)
        sendaddress(doomcom.numnodes).sin_addr = HostAddress(host)
        ' the first byte of the address is its lowest one
        If (sendaddress(doomcom.numnodes).sin_addr And &HFF&) <> 127 Then localonly = False

        doomcom.numnodes = doomcom.numnodes + 1
    Loop

    doomcom.id = DOOMCOM_ID
    doomcom.numplayers = doomcom.numnodes

    ' (the C did not check these)
    If doomcom.consoleplayer < 0 Or doomcom.consoleplayer >= doomcom.numplayers Then
        I_Error "-net: the player number must be 1 to " & doomcom.numplayers & _
            " (one more than the number of hosts)"
    End If

    ' build message to receive
    insocket = UDPsocket()
    BindToLocalPort insocket, DOOMPORT, localonly
    nb = 1
    ioctlsocket insocket, FIONBIO, nb

    I_Printf "I_InitNetwork: player " & (doomcom.consoleplayer + 1) & ", UDP port " & DOOMPORT & _
        ", " & (doomcom.numnodes - 1) & " other node(s)"
End Sub

Public Sub I_ShutdownNetwork()
    If insocket <> 0 Then
        closesocket insocket
        insocket = 0
    End If
    If wsaup Then
        WSACleanup
        wsaup = False
    End If
End Sub

Public Sub I_NetCmd()
    If doomcom.command = CMD_SEND Then
        PacketSend
    ElseIf doomcom.command = CMD_GET Then
        PacketGet
    Else
        I_Error "Bad net cmd: " & doomcom.command
    End If
End Sub
