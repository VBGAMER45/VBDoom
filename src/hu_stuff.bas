Attribute VB_Name = "hu_stuff"
' hu_stuff.bas - heads-up display: player messages, the map title on the
' automap and multiplayer chat (hu_stuff.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' Locally used constants, shortcuts.
Private Const HU_TITLEHEIGHT As Long = 1
Private Const HU_TITLEX As Long = 0
Private Const HU_INPUTTOGGLE As Long = 116       ' 't'
Private Const HU_INPUTWIDTH As Long = 64
Private Const HU_INPUTHEIGHT As Long = 1

' hu_stuff.h
Private Const HU_BROADCAST As Long = 5
Private Const HU_MSGREFRESH As Long = KEY_ENTER
Private Const HU_MSGX As Long = 0
Private Const HU_MSGY As Long = 0
Private Const HU_MSGWIDTH As Long = 64           ' in characters
Private Const HU_MSGHEIGHT As Long = 1           ' in lines
Private Const HU_MSGTIMEOUT As Long = 4 * TICRATE

Private Const QUEUESIZE As Long = 128

Private plr As Long                              ' player number (C: player_t* plr)
Private w_title As hu_textline_t
Private w_chat As hu_itext_t
Private always_off As Boolean
Private chat_dest(0 To 3) As Long
Private w_inputbuffer(0 To 3) As hu_itext_t

Private message_on As Boolean
Private message_nottobefuckedwith As Boolean

Private w_message As hu_stext_t
Private message_counter As Long

Private headsupactive As Boolean

' Builtin map names.  The actual names can be found in dstrings.
Private mapnames(0 To 44) As String              ' DOOM shareware/registered/retail (Ultimate) names
Private mapnames2(0 To 31) As String             ' DOOM 2 map names
Private mapnamesp(0 To 31) As String             ' Plutonia WAD map names
Private mapnamest(0 To 31) As String             ' TNT WAD map names

Private shiftxform(0 To 127) As Long
Private frenchKeyMap(0 To 127) As Long

Private chatchars(0 To 127) As Long              ' QUEUESIZE
Private head As Long
Private tail As Long

' HU_Responder statics
Private lastmessage As String
Private shiftdown As Boolean
Private altdown As Boolean
Private num_nobrainers As Long

' Fills a 128 entry character table: 0..31 and 127 map to themselves,
' chars holds the entries for 32..126.
Private Sub SetCharTable(tbl() As Long, ByVal chars As String)
    Dim i As Long
    For i = 0 To 31
        tbl(i) = i
    Next
    For i = 1 To Len(chars)
        tbl(31 + i) = Asc(Mid$(chars, i, 1))
    Next
    tbl(127) = 127
End Sub

Private Sub InitTables()
    Dim i As Long

    mapnames(0) = HUSTR_E1M1: mapnames(1) = HUSTR_E1M2: mapnames(2) = HUSTR_E1M3
    mapnames(3) = HUSTR_E1M4: mapnames(4) = HUSTR_E1M5: mapnames(5) = HUSTR_E1M6
    mapnames(6) = HUSTR_E1M7: mapnames(7) = HUSTR_E1M8: mapnames(8) = HUSTR_E1M9
    mapnames(9) = HUSTR_E2M1: mapnames(10) = HUSTR_E2M2: mapnames(11) = HUSTR_E2M3
    mapnames(12) = HUSTR_E2M4: mapnames(13) = HUSTR_E2M5: mapnames(14) = HUSTR_E2M6
    mapnames(15) = HUSTR_E2M7: mapnames(16) = HUSTR_E2M8: mapnames(17) = HUSTR_E2M9
    mapnames(18) = HUSTR_E3M1: mapnames(19) = HUSTR_E3M2: mapnames(20) = HUSTR_E3M3
    mapnames(21) = HUSTR_E3M4: mapnames(22) = HUSTR_E3M5: mapnames(23) = HUSTR_E3M6
    mapnames(24) = HUSTR_E3M7: mapnames(25) = HUSTR_E3M8: mapnames(26) = HUSTR_E3M9
    mapnames(27) = HUSTR_E4M1: mapnames(28) = HUSTR_E4M2: mapnames(29) = HUSTR_E4M3
    mapnames(30) = HUSTR_E4M4: mapnames(31) = HUSTR_E4M5: mapnames(32) = HUSTR_E4M6
    mapnames(33) = HUSTR_E4M7: mapnames(34) = HUSTR_E4M8: mapnames(35) = HUSTR_E4M9
    For i = 36 To 44
        mapnames(i) = "NEWLEVEL"
    Next

    mapnames2(0) = HUSTR_1: mapnames2(1) = HUSTR_2: mapnames2(2) = HUSTR_3: mapnames2(3) = HUSTR_4
    mapnames2(4) = HUSTR_5: mapnames2(5) = HUSTR_6: mapnames2(6) = HUSTR_7: mapnames2(7) = HUSTR_8
    mapnames2(8) = HUSTR_9: mapnames2(9) = HUSTR_10: mapnames2(10) = HUSTR_11: mapnames2(11) = HUSTR_12
    mapnames2(12) = HUSTR_13: mapnames2(13) = HUSTR_14: mapnames2(14) = HUSTR_15: mapnames2(15) = HUSTR_16
    mapnames2(16) = HUSTR_17: mapnames2(17) = HUSTR_18: mapnames2(18) = HUSTR_19: mapnames2(19) = HUSTR_20
    mapnames2(20) = HUSTR_21: mapnames2(21) = HUSTR_22: mapnames2(22) = HUSTR_23: mapnames2(23) = HUSTR_24
    mapnames2(24) = HUSTR_25: mapnames2(25) = HUSTR_26: mapnames2(26) = HUSTR_27: mapnames2(27) = HUSTR_28
    mapnames2(28) = HUSTR_29: mapnames2(29) = HUSTR_30: mapnames2(30) = HUSTR_31: mapnames2(31) = HUSTR_32

    mapnamesp(0) = PHUSTR_1: mapnamesp(1) = PHUSTR_2: mapnamesp(2) = PHUSTR_3: mapnamesp(3) = PHUSTR_4
    mapnamesp(4) = PHUSTR_5: mapnamesp(5) = PHUSTR_6: mapnamesp(6) = PHUSTR_7: mapnamesp(7) = PHUSTR_8
    mapnamesp(8) = PHUSTR_9: mapnamesp(9) = PHUSTR_10: mapnamesp(10) = PHUSTR_11: mapnamesp(11) = PHUSTR_12
    mapnamesp(12) = PHUSTR_13: mapnamesp(13) = PHUSTR_14: mapnamesp(14) = PHUSTR_15: mapnamesp(15) = PHUSTR_16
    mapnamesp(16) = PHUSTR_17: mapnamesp(17) = PHUSTR_18: mapnamesp(18) = PHUSTR_19: mapnamesp(19) = PHUSTR_20
    mapnamesp(20) = PHUSTR_21: mapnamesp(21) = PHUSTR_22: mapnamesp(22) = PHUSTR_23: mapnamesp(23) = PHUSTR_24
    mapnamesp(24) = PHUSTR_25: mapnamesp(25) = PHUSTR_26: mapnamesp(26) = PHUSTR_27: mapnamesp(27) = PHUSTR_28
    mapnamesp(28) = PHUSTR_29: mapnamesp(29) = PHUSTR_30: mapnamesp(30) = PHUSTR_31: mapnamesp(31) = PHUSTR_32

    mapnamest(0) = THUSTR_1: mapnamest(1) = THUSTR_2: mapnamest(2) = THUSTR_3: mapnamest(3) = THUSTR_4
    mapnamest(4) = THUSTR_5: mapnamest(5) = THUSTR_6: mapnamest(6) = THUSTR_7: mapnamest(7) = THUSTR_8
    mapnamest(8) = THUSTR_9: mapnamest(9) = THUSTR_10: mapnamest(10) = THUSTR_11: mapnamest(11) = THUSTR_12
    mapnamest(12) = THUSTR_13: mapnamest(13) = THUSTR_14: mapnamest(14) = THUSTR_15: mapnamest(15) = THUSTR_16
    mapnamest(16) = THUSTR_17: mapnamest(17) = THUSTR_18: mapnamest(18) = THUSTR_19: mapnamest(19) = THUSTR_20
    mapnamest(20) = THUSTR_21: mapnamest(21) = THUSTR_22: mapnamest(22) = THUSTR_23: mapnamest(23) = THUSTR_24
    mapnamest(24) = THUSTR_25: mapnamest(25) = THUSTR_26: mapnamest(26) = THUSTR_27: mapnamest(27) = THUSTR_28
    mapnamest(28) = THUSTR_29: mapnamest(29) = THUSTR_30: mapnamest(30) = THUSTR_31: mapnamest(31) = THUSTR_32

    ' Deviation: the C tests "if (french)", which is the enum constant and
    ' always true; the French tables are used for the French version only.
    If language = french Then
        SetCharTable shiftxform, " !""#$%&""()*+?_>?0123456789/.<+>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[!]""_'ABCDEFGHIJKLMNOPQRSTUVWXYZ{|}~"
    Else
        SetCharTable shiftxform, " !""#$%&""()*+<_>?)!@#$%^&*(::<+>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[!]""_'ABCDEFGHIJKLMNOPQRSTUVWXYZ{|}~"
    End If
    SetCharTable frenchKeyMap, " !""#$%&%()*+;-:!0123456789:M<=>?@QBCDEFGHIJKL,NOPARSTUVZXYW^\$^_@QBCDEFGHIJKL,NOPARSTUVZXYW^\$^"
End Sub

Private Function ForeignTranslation(ByVal ch As Long) As Long
    If ch < 128 Then ForeignTranslation = frenchKeyMap(ch) Else ForeignTranslation = ch
End Function

Public Sub HU_Init()
    Dim i As Long, j As Long

    InitTables

    ' load the heads-up font
    j = HU_FONTSTART
    For i = 0 To HU_FONTSIZE - 1
        hu_font(i) = W_GetNumForName("STCFN" & Format$(j, "000"))
        j = j + 1
    Next
End Sub

Public Sub HU_Stop()
    headsupactive = False
End Sub

Public Sub HU_Start()
    Dim i As Long, s As String, n As Long

    If headsupactive Then HU_Stop

    plr = consoleplayer
    message_on = False
    message_dontfuckwithme = False
    message_nottobefuckedwith = False
    chat_on = False

    ' create the message widget
    HUlib_initSText w_message, HU_MSGX, HU_MSGY, HU_MSGHEIGHT, HU_FONTSTART

    ' create the map title widget (HU_TITLEY = 167 - font height)
    HUlib_initTextLine w_title, HU_TITLEX, 167 - PatchHeight(hu_font(0)), HU_FONTSTART

    ' Deviation: the C left Plutonia and TNT on the DOOM 2 names (FIXME);
    ' gamemission tells them apart here.  Out of range maps get no title.
    Select Case gamemode
        Case shareware, registered, retail
            n = (gameepisode - 1) * 9 + gamemap - 1
            If n >= 0 And n <= 44 Then s = mapnames(n)
        Case Else
            n = gamemap - 1
            If n >= 0 And n <= 31 Then
                If gamemission = pack_plut Then
                    s = mapnamesp(n)
                ElseIf gamemission = pack_tnt Then
                    s = mapnamest(n)
                Else
                    s = mapnames2(n)
                End If
            End If
    End Select

    For i = 1 To Len(s)
        HUlib_addCharToTextLine w_title, Asc(Mid$(s, i, 1))
    Next

    ' create the chat widget (HU_INPUTX, HU_INPUTY)
    HUlib_initIText w_chat, HU_MSGX, HU_MSGY + HU_MSGHEIGHT * (PatchHeight(hu_font(0)) + 1), HU_FONTSTART

    ' create the inputbuffer widgets (switched on by always_off)
    For i = 0 To MAXPLAYERS - 1
        HUlib_initIText w_inputbuffer(i), 0, 0, 0
    Next

    headsupactive = True
End Sub

Public Sub HU_Drawer()
    HUlib_drawSText w_message, message_on
    HUlib_drawIText w_chat, chat_on
    If automapactive Then HUlib_drawTextLine w_title, False
End Sub

Public Sub HU_Erase()
    HUlib_eraseSText w_message, message_on
    HUlib_eraseIText w_chat, chat_on
    HUlib_eraseTextLine w_title
End Sub

Public Sub HU_Ticker()
    Dim i As Long, rc As Boolean, c As Long

    ' tick down message counter if message is up
    If message_counter <> 0 Then
        message_counter = message_counter - 1
        If message_counter = 0 Then
            message_on = False
            message_nottobefuckedwith = False
        End If
    End If

    If showMessages <> 0 Or message_dontfuckwithme Then
        ' display message if necessary
        If players(plr).message <> "" And (Not message_nottobefuckedwith Or message_dontfuckwithme) Then
            HUlib_addMessageToSText w_message, "", players(plr).message
            players(plr).message = ""
            message_on = True
            message_counter = HU_MSGTIMEOUT
            message_nottobefuckedwith = message_dontfuckwithme
            message_dontfuckwithme = False
        End If
    End If                                       ' else message_on = false;

    ' check for incoming chat characters
    If netgame Then
        For i = 0 To MAXPLAYERS - 1
            If playeringame(i) Then
                c = players(i).cmd.chatchar
                If i <> consoleplayer And c <> 0 Then
                    If c <= HU_BROADCAST Then
                        chat_dest(i) = c
                    Else
                        If c >= 97 And c <= 122 Then c = shiftxform(c)
                        rc = HUlib_keyInIText(w_inputbuffer(i), c)
                        If rc And c = KEY_ENTER Then
                            If Len(w_inputbuffer(i).l.txt) <> 0 And _
                               (chat_dest(i) = consoleplayer + 1 Or chat_dest(i) = HU_BROADCAST) Then
                                HUlib_addMessageToSText w_message, player_names(i), w_inputbuffer(i).l.txt

                                message_nottobefuckedwith = True
                                message_on = True
                                message_counter = HU_MSGTIMEOUT
                                If gamemode = commercial Then
                                    S_StartSound 0, sfx_radio
                                Else
                                    S_StartSound 0, sfx_tink
                                End If
                            End If
                            HUlib_resetIText w_inputbuffer(i)
                        End If
                    End If
                    players(i).cmd.chatchar = 0
                End If
            End If
        Next
    End If
End Sub

Public Sub HU_queueChatChar(ByVal c As Long)
    If ((head + 1) And (QUEUESIZE - 1)) = tail Then
        players(plr).message = HUSTR_MSGU
    Else
        chatchars(head) = c
        head = (head + 1) And (QUEUESIZE - 1)
    End If
End Sub

Public Function HU_dequeueChatChar() As Long
    If head <> tail Then
        HU_dequeueChatChar = chatchars(tail)
        tail = (tail + 1) And (QUEUESIZE - 1)
    Else
        HU_dequeueChatChar = 0
    End If
End Function

Public Function HU_Responder(ev As event_t) As Boolean
    Dim macromessage As String, eatkey As Boolean, c As Long
    Dim i As Long, numplayers As Long
    Dim destination_keys(0 To 3) As Long

    destination_keys(0) = HUSTR_KEYGREEN
    destination_keys(1) = HUSTR_KEYINDIGO
    destination_keys(2) = HUSTR_KEYBROWN
    destination_keys(3) = HUSTR_KEYRED

    numplayers = 0
    For i = 0 To MAXPLAYERS - 1
        If playeringame(i) Then numplayers = numplayers + 1
    Next

    If ev.data1 = KEY_RSHIFT Then
        shiftdown = (ev.typ = ev_keydown)
        HU_Responder = False
        Exit Function
    ElseIf ev.data1 = KEY_RALT Or ev.data1 = KEY_LALT Then
        altdown = (ev.typ = ev_keydown)
        HU_Responder = False
        Exit Function
    End If

    If ev.typ <> ev_keydown Then
        HU_Responder = False
        Exit Function
    End If

    If Not chat_on Then
        If ev.data1 = HU_MSGREFRESH Then
            message_on = True
            message_counter = HU_MSGTIMEOUT
            eatkey = True
        ElseIf netgame And ev.data1 = HU_INPUTTOGGLE Then
            eatkey = True
            chat_on = True
            HUlib_resetIText w_chat
            HU_queueChatChar HU_BROADCAST
        ElseIf netgame And numplayers > 2 Then
            For i = 0 To MAXPLAYERS - 1
                If ev.data1 = destination_keys(i) Then
                    If playeringame(i) And i <> consoleplayer Then
                        eatkey = True
                        chat_on = True
                        HUlib_resetIText w_chat
                        HU_queueChatChar i + 1
                        Exit For
                    ElseIf i = consoleplayer Then
                        num_nobrainers = num_nobrainers + 1
                        If num_nobrainers < 3 Then
                            players(plr).message = HUSTR_TALKTOSELF1
                        ElseIf num_nobrainers < 6 Then
                            players(plr).message = HUSTR_TALKTOSELF2
                        ElseIf num_nobrainers < 9 Then
                            players(plr).message = HUSTR_TALKTOSELF3
                        ElseIf num_nobrainers < 32 Then
                            players(plr).message = HUSTR_TALKTOSELF4
                        Else
                            players(plr).message = HUSTR_TALKTOSELF5
                        End If
                    End If
                End If
            Next
        End If
    Else
        c = ev.data1 And 255                     ' unsigned char
        ' send a macro
        If altdown Then
            c = c - 48                           ' '0'
            If c < 0 Or c > 9 Then
                HU_Responder = False
                Exit Function
            End If
            macromessage = chat_macros(c)

            ' kill last message with a '\n'
            HU_queueChatChar KEY_ENTER           ' DEBUG!!!

            ' send the macro message
            For i = 1 To Len(macromessage)
                HU_queueChatChar Asc(Mid$(macromessage, i, 1))
            Next
            HU_queueChatChar KEY_ENTER

            ' leave chat mode and notify that it was sent
            chat_on = False
            lastmessage = chat_macros(c)
            players(plr).message = lastmessage
            eatkey = True
        Else
            If language = french Then c = ForeignTranslation(c)
            If shiftdown Or (c >= 97 And c <= 122) Then
                If c < 128 Then c = shiftxform(c) ' the C reads past the table otherwise
            End If
            eatkey = HUlib_keyInIText(w_chat, c)
            If eatkey Then HU_queueChatChar c
            If c = KEY_ENTER Then
                chat_on = False
                If Len(w_chat.l.txt) <> 0 Then
                    lastmessage = w_chat.l.txt
                    players(plr).message = lastmessage
                End If
            ElseIf c = KEY_ESCAPE Then
                chat_on = False
            End If
        End If
    End If

    HU_Responder = eatkey
End Function
