Attribute VB_Name = "am_map"
' am_map.bas - the automap code (am_map.c).
'
' automapactive is declared in doomdef, the AM_MSG* constants as well.
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' For use if I do walls with outsides/insides
Private Const REDS As Long = 256 - 5 * 16
Private Const REDRANGE As Long = 16
Private Const BLUES As Long = 256 - 4 * 16 + 8
Private Const BLUERANGE As Long = 8
Private Const GREENS As Long = 7 * 16
Private Const GREENRANGE As Long = 16
Private Const GRAYS As Long = 6 * 16
Private Const GRAYSRANGE As Long = 16
Private Const BROWNS As Long = 4 * 16
Private Const BROWNRANGE As Long = 16
Private Const YELLOWS As Long = 256 - 32 + 7
Private Const YELLOWRANGE As Long = 1
Private Const BLACK As Long = 0
Private Const WHITE As Long = 256 - 47

' Automap colors
Private Const BACKGROUND As Long = BLACK
Private Const YOURCOLORS As Long = WHITE
Private Const YOURRANGE As Long = 0
Private Const WALLCOLORS As Long = REDS
Private Const WALLRANGE As Long = REDRANGE
Private Const TSWALLCOLORS As Long = GRAYS
Private Const TSWALLRANGE As Long = GRAYSRANGE
Private Const FDWALLCOLORS As Long = BROWNS
Private Const FDWALLRANGE As Long = BROWNRANGE
Private Const CDWALLCOLORS As Long = YELLOWS
Private Const CDWALLRANGE As Long = YELLOWRANGE
Private Const THINGCOLORS As Long = GREENS
Private Const THINGRANGE As Long = GREENRANGE
Private Const SECRETWALLCOLORS As Long = WALLCOLORS
Private Const SECRETWALLRANGE As Long = WALLRANGE
Private Const GRIDCOLORS As Long = GRAYS + GRAYSRANGE \ 2
Private Const GRIDRANGE As Long = 0
Private Const XHAIRCOLORS As Long = GRAYS

' drawing stuff
Private Const FB As Long = 0

Private Const AM_PANDOWNKEY As Long = KEY_DOWNARROW
Private Const AM_PANUPKEY As Long = KEY_UPARROW
Private Const AM_PANRIGHTKEY As Long = KEY_RIGHTARROW
Private Const AM_PANLEFTKEY As Long = KEY_LEFTARROW
Private Const AM_ZOOMINKEY As Long = 61          ' '='
Private Const AM_ZOOMOUTKEY As Long = 45         ' '-'
Private Const AM_STARTKEY As Long = KEY_TAB
Private Const AM_ENDKEY As Long = KEY_TAB
Private Const AM_GOBIGKEY As Long = 48           ' '0'
Private Const AM_FOLLOWKEY As Long = 102         ' 'f'
Private Const AM_GRIDKEY As Long = 103           ' 'g'
Private Const AM_MARKKEY As Long = 109           ' 'm'
Private Const AM_CLEARMARKKEY As Long = 99       ' 'c'

Private Const AM_NUMMARKPOINTS As Long = 10

' scale on entry (C: .2*FRACUNIT, truncated by the initializer)
Private Const INITSCALEMTOF As Long = 13107
' how much the automap moves window per tic in frame-buffer coordinates
' moves 140 pixels in 1 second
Private Const F_PANINC As Long = 4
' how much zoom-in per tic
' goes to 2x in 1 second
Private Const M_ZOOMIN As Long = 66846           ' (int) (1.02*FRACUNIT)
' how much zoom-out per tic
' pulls out to 0.5x in 1 second
Private Const M_ZOOMOUT As Long = 64250          ' (int) (FRACUNIT/1.02)

' the following is crap
Private Const LINE_NEVERSEE As Long = ML_DONTDRAW

Private Type fpoint_t
    x As Long
    y As Long
End Type

Private Type fline_t
    a As fpoint_t
    b As fpoint_t
End Type

Private Type mpoint_t
    x As Long
    y As Long
End Type

Private Type mline_t
    a As mpoint_t
    b As mpoint_t
End Type

' The vector graphics for the automap.
Private player_arrow(0 To 6) As mline_t
Private cheat_player_arrow(0 To 15) As mline_t
Private triangle_guy(0 To 2) As mline_t
Private thintriangle_guy(0 To 2) As mline_t
Private shapesInit As Boolean

Private cheating As Long
Private grid As Boolean

Private leveljuststarted As Long                 ' kluge until AM_LevelInit() is called

Private Const finit_width As Long = SCREENWIDTH
Private Const finit_height As Long = SCREENHEIGHT - 32

' location of window on screen
Private f_x As Long
Private f_y As Long

' size of window on screen
Private f_w As Long
Private f_h As Long

Private lightlev As Long                         ' used for funky strobing effect
Private amclock As Long

Private m_paninc As mpoint_t                     ' how far the window pans each tic (map coords)
Private mtof_zoommul As Long                     ' how far the window zooms in each tic (map coords)
Private ftom_zoommul As Long                     ' how far the window zooms in each tic (fb coords)

Private m_x As Long                              ' LL x,y where the window is on the map (map coords)
Private m_y As Long
Private m_x2 As Long                             ' UR x,y where the window is on the map (map coords)
Private m_y2 As Long

' width/height of window on map (map coords)
Private m_w As Long
Private m_h As Long

' based on level size
Private min_x As Long
Private min_y As Long
Private max_x As Long
Private max_y As Long

Private max_w As Long                            ' max_x-min_x,
Private max_h As Long                            ' max_y-min_y

' based on player size
Private min_w As Long
Private min_h As Long

Private min_scale_mtof As Long                   ' used to tell when to stop zooming out
Private max_scale_mtof As Long                   ' used to tell when to stop zooming in

' old stuff for recovery later
Private old_m_w As Long
Private old_m_h As Long
Private old_m_x As Long
Private old_m_y As Long

' old location used by the Follower routine
Private f_oldloc As mpoint_t

' used by MTOF to scale from map-to-frame-buffer coords
Private scale_mtof As Long
' used by FTOM to scale from frame-buffer-to-map coords (=1/scale_mtof)
Private scale_ftom As Long

Private plr As Long                              ' the player represented by an arrow (C: player_t*)

Private marknums(0 To 9) As Long                 ' numbers used for marking by the automap
Private markpoints(0 To 9) As mpoint_t           ' where the points are
Private markpointnum As Long                     ' next point to be assigned

Private followplayer As Boolean                  ' specifies whether to follow the player around

Private cheat_amap As cheatseq_t

Private stopped As Boolean

' AM_Start statics
Private lastlevel As Long
Private lastepisode As Long

' AM_Responder statics
Private cheatstate As Long
Private bigstate As Boolean

' one-time setup of the C static initializers
Private Sub AM_InitStatics()
    Dim r As Long

    If shapesInit Then Exit Sub
    shapesInit = True

    leveljuststarted = 1
    scale_mtof = INITSCALEMTOF
    followplayer = True
    stopped = True
    lastlevel = -1
    lastepisode = -1
    cht_MakeSeq cheat_amap, "iddt"

    ' A line drawing of the player pointing right, starting from the middle.
    r = (8 * PLAYERRADIUS) \ 7
    SetML player_arrow(0), -r + r \ 8, 0, r, 0                  ' -----
    SetML player_arrow(1), r, 0, r - r \ 2, r \ 4               ' ----->
    SetML player_arrow(2), r, 0, r - r \ 2, -r \ 4
    SetML player_arrow(3), -r + r \ 8, 0, -r - r \ 8, r \ 4     ' >---->
    SetML player_arrow(4), -r + r \ 8, 0, -r - r \ 8, -r \ 4
    SetML player_arrow(5), -r + 3 * r \ 8, 0, -r + r \ 8, r \ 4 ' >>--->
    SetML player_arrow(6), -r + 3 * r \ 8, 0, -r + r \ 8, -r \ 4

    SetML cheat_player_arrow(0), -r + r \ 8, 0, r, 0                  ' -----
    SetML cheat_player_arrow(1), r, 0, r - r \ 2, r \ 6               ' ----->
    SetML cheat_player_arrow(2), r, 0, r - r \ 2, -r \ 6
    SetML cheat_player_arrow(3), -r + r \ 8, 0, -r - r \ 8, r \ 6     ' >----->
    SetML cheat_player_arrow(4), -r + r \ 8, 0, -r - r \ 8, -r \ 6
    SetML cheat_player_arrow(5), -r + 3 * r \ 8, 0, -r + r \ 8, r \ 6 ' >>----->
    SetML cheat_player_arrow(6), -r + 3 * r \ 8, 0, -r + r \ 8, -r \ 6
    SetML cheat_player_arrow(7), -r \ 2, 0, -r \ 2, -r \ 6            ' >>-d--->
    SetML cheat_player_arrow(8), -r \ 2, -r \ 6, -r \ 2 + r \ 6, -r \ 6
    SetML cheat_player_arrow(9), -r \ 2 + r \ 6, -r \ 6, -r \ 2 + r \ 6, r \ 4
    SetML cheat_player_arrow(10), -r \ 6, 0, -r \ 6, -r \ 6           ' >>-dd-->
    SetML cheat_player_arrow(11), -r \ 6, -r \ 6, 0, -r \ 6
    SetML cheat_player_arrow(12), 0, -r \ 6, 0, r \ 4
    SetML cheat_player_arrow(13), r \ 6, r \ 4, r \ 6, -r \ 7         ' >>-ddt->
    SetML cheat_player_arrow(14), r \ 6, -r \ 7, r \ 6 + r \ 32, -r \ 7 - r \ 32
    SetML cheat_player_arrow(15), r \ 6 + r \ 32, -r \ 7 - r \ 32, r \ 6 + r \ 10, -r \ 7

    ' (C: -.867*R and so on, truncated by the initializers)
    SetML triangle_guy(0), -56819, -32768, 56819, -32768
    SetML triangle_guy(1), 56819, -32768, 0, FRACUNIT
    SetML triangle_guy(2), 0, FRACUNIT, -56819, -32768

    SetML thintriangle_guy(0), -32768, -45875, FRACUNIT, 0
    SetML thintriangle_guy(1), FRACUNIT, 0, -32768, 45875
    SetML thintriangle_guy(2), -32768, 45875, -32768, -45875
End Sub

Private Sub SetML(ml As mline_t, ByVal ax As Long, ByVal ay As Long, ByVal bx As Long, ByVal by As Long)
    ml.a.x = ax
    ml.a.y = ay
    ml.b.x = bx
    ml.b.y = by
End Sub

' translates between frame-buffer and map distances
Private Function FTOM(ByVal x As Long) As Long
    FTOM = FixedMul(SHL(x, 16), scale_ftom)
End Function

Private Function MTOF(ByVal x As Long) As Long
    MTOF = SAR(FixedMul(x, scale_mtof), 16)
End Function

' translates between frame-buffer and map coordinates
Private Function CXMTOF(ByVal x As Long) As Long
    CXMTOF = f_x + MTOF(x - m_x)
End Function

Private Function CYMTOF(ByVal y As Long) As Long
    CYMTOF = f_y + (f_h - MTOF(y - m_y))
End Function

Private Sub AM_activateNewScale()
    m_x = m_x + m_w \ 2
    m_y = m_y + m_h \ 2
    m_w = FTOM(f_w)
    m_h = FTOM(f_h)
    m_x = m_x - m_w \ 2
    m_y = m_y - m_h \ 2
    m_x2 = m_x + m_w
    m_y2 = m_y + m_h
End Sub

Private Sub AM_saveScaleAndLoc()
    old_m_x = m_x
    old_m_y = m_y
    old_m_w = m_w
    old_m_h = m_h
End Sub

Private Sub AM_restoreScaleAndLoc()
    m_w = old_m_w
    m_h = old_m_h
    If Not followplayer Then
        m_x = old_m_x
        m_y = old_m_y
    Else
        m_x = mobjs(players(plr).mo).x - m_w \ 2
        m_y = mobjs(players(plr).mo).y - m_h \ 2
    End If
    m_x2 = m_x + m_w
    m_y2 = m_y + m_h

    ' Change the scaling multipliers
    scale_mtof = FixedDiv(SHL(f_w, FRACBITS), m_w)
    scale_ftom = FixedDiv(FRACUNIT, scale_mtof)
End Sub

' adds a marker at the current location
Private Sub AM_addMark()
    markpoints(markpointnum).x = m_x + m_w \ 2
    markpoints(markpointnum).y = m_y + m_h \ 2
    markpointnum = (markpointnum + 1) Mod AM_NUMMARKPOINTS
End Sub

' Determines bounding box of all vertices, sets global variables
' controlling zoom range.
Private Sub AM_findMinMaxBoundaries()
    Dim i As Long, a As Long, b As Long

    min_x = MAXINT: min_y = MAXINT
    max_x = -MAXINT: max_y = -MAXINT

    For i = 0 To numvertexes - 1
        If vertexes(i).x < min_x Then
            min_x = vertexes(i).x
        ElseIf vertexes(i).x > max_x Then
            max_x = vertexes(i).x
        End If

        If vertexes(i).y < min_y Then
            min_y = vertexes(i).y
        ElseIf vertexes(i).y > max_y Then
            max_y = vertexes(i).y
        End If
    Next

    max_w = max_x - min_x
    max_h = max_y - min_y

    min_w = 2 * PLAYERRADIUS                     ' const? never changed?
    min_h = 2 * PLAYERRADIUS

    a = FixedDiv(SHL(f_w, FRACBITS), max_w)
    b = FixedDiv(SHL(f_h, FRACBITS), max_h)

    If a < b Then min_scale_mtof = a Else min_scale_mtof = b
    max_scale_mtof = FixedDiv(SHL(f_h, FRACBITS), 2 * PLAYERRADIUS)
End Sub

Private Sub AM_changeWindowLoc()
    If m_paninc.x <> 0 Or m_paninc.y <> 0 Then
        followplayer = False
        f_oldloc.x = MAXINT
    End If

    m_x = m_x + m_paninc.x
    m_y = m_y + m_paninc.y

    If m_x + m_w \ 2 > max_x Then
        m_x = max_x - m_w \ 2
    ElseIf m_x + m_w \ 2 < min_x Then
        m_x = min_x - m_w \ 2
    End If

    If m_y + m_h \ 2 > max_y Then
        m_y = max_y - m_h \ 2
    ElseIf m_y + m_h \ 2 < min_y Then
        m_y = min_y - m_h \ 2
    End If

    m_x2 = m_x + m_w
    m_y2 = m_y + m_h
End Sub

Private Sub AM_initVariables()
    Dim pnum As Long, st_notify As event_t

    automapactive = True

    f_oldloc.x = MAXINT
    amclock = 0
    lightlev = 0

    m_paninc.x = 0
    m_paninc.y = 0
    ftom_zoommul = FRACUNIT
    mtof_zoommul = FRACUNIT

    m_w = FTOM(f_w)
    m_h = FTOM(f_h)

    ' find player to center on initially
    pnum = consoleplayer
    If Not playeringame(pnum) Then
        For pnum = 0 To MAXPLAYERS - 1
            If playeringame(pnum) Then Exit For
        Next
    End If

    plr = pnum
    m_x = mobjs(players(plr).mo).x - m_w \ 2
    m_y = mobjs(players(plr).mo).y - m_h \ 2
    AM_changeWindowLoc

    ' for saving & restoring
    old_m_x = m_x
    old_m_y = m_y
    old_m_w = m_w
    old_m_h = m_h

    ' inform the status bar of the change
    st_notify.typ = ev_keyup
    st_notify.data1 = AM_MSGENTERED
    ST_Responder st_notify
End Sub

Private Sub AM_loadPics()
    Dim i As Long
    For i = 0 To 9
        marknums(i) = W_GetNumForName("AMMNUM" & i)
    Next
End Sub

Private Sub AM_clearMarks()
    Dim i As Long
    For i = 0 To AM_NUMMARKPOINTS - 1
        markpoints(i).x = -1                     ' means empty
    Next
    markpointnum = 0
End Sub

' should be called at the start of every level
' right now, i figure it out myself
Private Sub AM_LevelInit()
    leveljuststarted = 0

    f_x = 0
    f_y = 0
    f_w = finit_width
    f_h = finit_height

    AM_clearMarks

    AM_findMinMaxBoundaries
    scale_mtof = FixedDiv(min_scale_mtof, 45875)  ' (int) (0.7*FRACUNIT)
    If scale_mtof > max_scale_mtof Then scale_mtof = min_scale_mtof
    scale_ftom = FixedDiv(FRACUNIT, scale_mtof)
End Sub

' Deviation: the C initialised its notify event as { 0, ev_keyup,
' AM_MSGEXITED }, so the status bar got a keydown of key 1 instead.
Public Sub AM_Stop()
    Dim st_notify As event_t

    AM_InitStatics
    automapactive = False
    st_notify.typ = ev_keyup
    st_notify.data1 = AM_MSGEXITED
    ST_Responder st_notify
    stopped = True
End Sub

Public Sub AM_Start()
    AM_InitStatics

    If Not stopped Then AM_Stop
    stopped = False
    If lastlevel <> gamemap Or lastepisode <> gameepisode Then
        AM_LevelInit
        lastlevel = gamemap
        lastepisode = gameepisode
    End If
    AM_initVariables
    AM_loadPics
End Sub

' set the window scale to the maximum size
Private Sub AM_minOutWindowScale()
    scale_mtof = min_scale_mtof
    scale_ftom = FixedDiv(FRACUNIT, scale_mtof)
    AM_activateNewScale
End Sub

' set the window scale to the minimum size
Private Sub AM_maxOutWindowScale()
    scale_mtof = max_scale_mtof
    scale_ftom = FixedDiv(FRACUNIT, scale_mtof)
    AM_activateNewScale
End Sub

' Handle events (user inputs) in automap mode
Public Function AM_Responder(ev As event_t) As Boolean
    Dim rc As Boolean

    AM_InitStatics
    rc = False

    If Not automapactive Then
        If ev.typ = ev_keydown And ev.data1 = AM_STARTKEY Then
            AM_Start
            viewactive = False
            rc = True
        End If

    ElseIf ev.typ = ev_keydown Then
        rc = True
        Select Case ev.data1
            Case AM_PANRIGHTKEY                  ' pan right
                If Not followplayer Then m_paninc.x = FTOM(F_PANINC) Else rc = False
            Case AM_PANLEFTKEY                   ' pan left
                If Not followplayer Then m_paninc.x = -FTOM(F_PANINC) Else rc = False
            Case AM_PANUPKEY                     ' pan up
                If Not followplayer Then m_paninc.y = FTOM(F_PANINC) Else rc = False
            Case AM_PANDOWNKEY                   ' pan down
                If Not followplayer Then m_paninc.y = -FTOM(F_PANINC) Else rc = False
            Case AM_ZOOMOUTKEY                   ' zoom out
                mtof_zoommul = M_ZOOMOUT
                ftom_zoommul = M_ZOOMIN
            Case AM_ZOOMINKEY                    ' zoom in
                mtof_zoommul = M_ZOOMIN
                ftom_zoommul = M_ZOOMOUT
            Case AM_ENDKEY
                bigstate = False
                viewactive = True
                AM_Stop
            Case AM_GOBIGKEY
                bigstate = Not bigstate
                If bigstate Then
                    AM_saveScaleAndLoc
                    AM_minOutWindowScale
                Else
                    AM_restoreScaleAndLoc
                End If
            Case AM_FOLLOWKEY
                followplayer = Not followplayer
                f_oldloc.x = MAXINT
                If followplayer Then
                    players(plr).message = AMSTR_FOLLOWON
                Else
                    players(plr).message = AMSTR_FOLLOWOFF
                End If
            Case AM_GRIDKEY
                grid = Not grid
                If grid Then
                    players(plr).message = AMSTR_GRIDON
                Else
                    players(plr).message = AMSTR_GRIDOFF
                End If
            Case AM_MARKKEY
                players(plr).message = AMSTR_MARKEDSPOT & " " & markpointnum
                AM_addMark
            Case AM_CLEARMARKKEY
                AM_clearMarks
                players(plr).message = AMSTR_MARKSCLEARED
            Case Else
                cheatstate = 0
                rc = False
        End Select
        If deathmatch = 0 Then
            If cht_CheckCheat(cheat_amap, ev.data1) Then
                rc = False
                cheating = (cheating + 1) Mod 3
            End If
        End If

    ElseIf ev.typ = ev_keyup Then
        rc = False
        Select Case ev.data1
            Case AM_PANRIGHTKEY
                If Not followplayer Then m_paninc.x = 0
            Case AM_PANLEFTKEY
                If Not followplayer Then m_paninc.x = 0
            Case AM_PANUPKEY
                If Not followplayer Then m_paninc.y = 0
            Case AM_PANDOWNKEY
                If Not followplayer Then m_paninc.y = 0
            Case AM_ZOOMOUTKEY, AM_ZOOMINKEY
                mtof_zoommul = FRACUNIT
                ftom_zoommul = FRACUNIT
        End Select
    End If

    AM_Responder = rc
End Function

' Zooming
Private Sub AM_changeWindowScale()
    ' Change the scaling multipliers
    scale_mtof = FixedMul(scale_mtof, mtof_zoommul)
    scale_ftom = FixedDiv(FRACUNIT, scale_mtof)

    If scale_mtof < min_scale_mtof Then
        AM_minOutWindowScale
    ElseIf scale_mtof > max_scale_mtof Then
        AM_maxOutWindowScale
    Else
        AM_activateNewScale
    End If
End Sub

Private Sub AM_doFollowPlayer()
    Dim mo As Long
    mo = players(plr).mo
    If f_oldloc.x <> mobjs(mo).x Or f_oldloc.y <> mobjs(mo).y Then
        m_x = FTOM(MTOF(mobjs(mo).x)) - m_w \ 2
        m_y = FTOM(MTOF(mobjs(mo).y)) - m_h \ 2
        m_x2 = m_x + m_w
        m_y2 = m_y + m_h
        f_oldloc.x = mobjs(mo).x
        f_oldloc.y = mobjs(mo).y
    End If
End Sub

' (AM_updateLightLev is not called by the C code either.)

' Updates on Game Tick
Public Sub AM_Ticker()
    If Not automapactive Then Exit Sub

    amclock = amclock + 1

    If followplayer Then AM_doFollowPlayer

    ' Change the zoom if necessary
    If ftom_zoommul <> FRACUNIT Then AM_changeWindowScale

    ' Change x,y location
    If m_paninc.x <> 0 Or m_paninc.y <> 0 Then AM_changeWindowLoc

    ' Update light level
    ' AM_updateLightLev();
End Sub

' Clear automap frame buffer.
Private Sub AM_clearFB(ByVal color As Long)
    FillMemory screens(0), f_w * f_h, color
End Sub

' outcode of a frame-buffer point (C: DOOUTCODE)
Private Function DoOutcode(ByVal mx As Long, ByVal my As Long) As Long
    Dim oc As Long
    oc = 0
    If my < 0 Then
        oc = oc Or 8                             ' TOP
    ElseIf my >= f_h Then
        oc = oc Or 4                             ' BOTTOM
    End If
    If mx < 0 Then
        oc = oc Or 1                             ' LEFT
    ElseIf mx >= f_w Then
        oc = oc Or 2                             ' RIGHT
    End If
    DoOutcode = oc
End Function

' Automap clipping of lines.
'
' Based on Cohen-Sutherland clipping algorithm but with a slightly faster
' reject and precalculated slopes.  If the speed is needed, use a hash
' algorithm to handle the common cases.
Private Function AM_clipMline(ml As mline_t, fl As fline_t) As Boolean
    Const LEFT_ As Long = 1, RIGHT_ As Long = 2, BOTTOM_ As Long = 4, TOP_ As Long = 8
    Dim outcode1 As Long, outcode2 As Long, outside As Long
    Dim tmp As fpoint_t, dx As Long, dy As Long

    ' do trivial rejects and outcodes
    If ml.a.y > m_y2 Then
        outcode1 = TOP_
    ElseIf ml.a.y < m_y Then
        outcode1 = BOTTOM_
    End If

    If ml.b.y > m_y2 Then
        outcode2 = TOP_
    ElseIf ml.b.y < m_y Then
        outcode2 = BOTTOM_
    End If

    If (outcode1 And outcode2) <> 0 Then Exit Function ' trivially outside

    If ml.a.x < m_x Then
        outcode1 = outcode1 Or LEFT_
    ElseIf ml.a.x > m_x2 Then
        outcode1 = outcode1 Or RIGHT_
    End If

    If ml.b.x < m_x Then
        outcode2 = outcode2 Or LEFT_
    ElseIf ml.b.x > m_x2 Then
        outcode2 = outcode2 Or RIGHT_
    End If

    If (outcode1 And outcode2) <> 0 Then Exit Function ' trivially outside

    ' transform to frame-buffer coordinates.
    fl.a.x = CXMTOF(ml.a.x)
    fl.a.y = CYMTOF(ml.a.y)
    fl.b.x = CXMTOF(ml.b.x)
    fl.b.y = CYMTOF(ml.b.y)

    outcode1 = DoOutcode(fl.a.x, fl.a.y)
    outcode2 = DoOutcode(fl.b.x, fl.b.y)

    If (outcode1 And outcode2) <> 0 Then Exit Function

    Do While (outcode1 Or outcode2) <> 0
        ' may be partially inside box
        ' find an outside point
        If outcode1 <> 0 Then outside = outcode1 Else outside = outcode2

        ' clip to each side
        If (outside And TOP_) <> 0 Then
            dy = fl.a.y - fl.b.y
            dx = fl.b.x - fl.a.x
            tmp.x = fl.a.x + (dx * fl.a.y) \ dy
            tmp.y = 0
        ElseIf (outside And BOTTOM_) <> 0 Then
            dy = fl.a.y - fl.b.y
            dx = fl.b.x - fl.a.x
            tmp.x = fl.a.x + (dx * (fl.a.y - f_h)) \ dy
            tmp.y = f_h - 1
        ElseIf (outside And RIGHT_) <> 0 Then
            dy = fl.b.y - fl.a.y
            dx = fl.b.x - fl.a.x
            tmp.y = fl.a.y + (dy * (f_w - 1 - fl.a.x)) \ dx
            tmp.x = f_w - 1
        ElseIf (outside And LEFT_) <> 0 Then
            dy = fl.b.y - fl.a.y
            dx = fl.b.x - fl.a.x
            tmp.y = fl.a.y + (dy * (-fl.a.x)) \ dx
            tmp.x = 0
        End If

        If outside = outcode1 Then
            fl.a = tmp
            outcode1 = DoOutcode(fl.a.x, fl.a.y)
        Else
            fl.b = tmp
            outcode2 = DoOutcode(fl.b.x, fl.b.y)
        End If

        If (outcode1 And outcode2) <> 0 Then Exit Function ' trivially outside
    Loop

    AM_clipMline = True
End Function

' Classic Bresenham w/ whatever optimizations needed for speed
Private Sub AM_drawFline(fl As fline_t, ByVal color As Long)
    Dim x As Long, y As Long, dx As Long, dy As Long, sx As Long, sy As Long
    Dim ax As Long, ay As Long, d As Long

    ' For debugging only (the C printed "fuck %d" here)
    If fl.a.x < 0 Or fl.a.x >= f_w Or fl.a.y < 0 Or fl.a.y >= f_h _
       Or fl.b.x < 0 Or fl.b.x >= f_w Or fl.b.y < 0 Or fl.b.y >= f_h Then
        Exit Sub
    End If

    dx = fl.b.x - fl.a.x
    ax = 2 * Abs(dx)
    If dx < 0 Then sx = -1 Else sx = 1

    dy = fl.b.y - fl.a.y
    ay = 2 * Abs(dy)
    If dy < 0 Then sy = -1 Else sy = 1

    x = fl.a.x
    y = fl.a.y

    If ax > ay Then
        d = ay - ax \ 2
        Do
            screens(y * f_w + x) = color         ' PUTDOT
            If x = fl.b.x Then Exit Sub
            If d >= 0 Then
                y = y + sy
                d = d - ax
            End If
            x = x + sx
            d = d + ay
        Loop
    Else
        d = ax - ay \ 2
        Do
            screens(y * f_w + x) = color         ' PUTDOT
            If y = fl.b.y Then Exit Sub
            If d >= 0 Then
                x = x + sx
                d = d - ay
            End If
            y = y + sy
            d = d + ax
        Loop
    End If
End Sub

' Clip lines, draw visible part sof lines.
Private Sub AM_drawMline(ml As mline_t, ByVal color As Long)
    Dim fl As fline_t
    If AM_clipMline(ml, fl) Then AM_drawFline fl, color ' draws it on frame buffer using fb coords
End Sub

' Draws flat (floor/ceiling tile) aligned grid lines.
Private Sub AM_drawGrid(ByVal color As Long)
    Dim x As Long, y As Long, start As Long, finish As Long, ml As mline_t
    Const BLOCKSIZE As Long = MAPBLOCKUNITS * FRACUNIT ' MAPBLOCKUNITS<<FRACBITS

    ' Figure out start of vertical gridlines
    start = m_x
    If ((start - bmaporgx) Mod BLOCKSIZE) <> 0 Then
        start = start + BLOCKSIZE - ((start - bmaporgx) Mod BLOCKSIZE)
    End If
    finish = m_x + m_w

    ' draw vertical gridlines
    ml.a.y = m_y
    ml.b.y = m_y + m_h
    x = start
    Do While x < finish
        ml.a.x = x
        ml.b.x = x
        AM_drawMline ml, color
        x = x + BLOCKSIZE
    Loop

    ' Figure out start of horizontal gridlines
    start = m_y
    If ((start - bmaporgy) Mod BLOCKSIZE) <> 0 Then
        start = start + BLOCKSIZE - ((start - bmaporgy) Mod BLOCKSIZE)
    End If
    finish = m_y + m_h

    ' draw horizontal gridlines
    ml.a.x = m_x
    ml.b.x = m_x + m_w
    y = start
    Do While y < finish
        ml.a.y = y
        ml.b.y = y
        AM_drawMline ml, color
        y = y + BLOCKSIZE
    Loop
End Sub

' Determines visible lines, draws them.
' This is LineDef based, not LineSeg based.
Private Sub AM_drawWalls()
    Dim i As Long, l As mline_t, flags As Long, back As Long, front As Long

    For i = 0 To numlines - 1
        l.a.x = vertexes(lines(i).v1).x
        l.a.y = vertexes(lines(i).v1).y
        l.b.x = vertexes(lines(i).v2).x
        l.b.y = vertexes(lines(i).v2).y
        flags = lines(i).flags
        If cheating <> 0 Or (flags And ML_MAPPED) <> 0 Then
            If (flags And LINE_NEVERSEE) <> 0 And cheating = 0 Then GoTo nextline
            back = lines(i).backsector
            front = lines(i).frontsector
            If back = -1 Then
                AM_drawMline l, WALLCOLORS + lightlev
            Else
                If lines(i).special = 39 Then
                    ' teleporters
                    AM_drawMline l, WALLCOLORS + WALLRANGE \ 2
                ElseIf (flags And ML_SECRET) <> 0 Then
                    ' secret door
                    If cheating <> 0 Then
                        AM_drawMline l, SECRETWALLCOLORS + lightlev
                    Else
                        AM_drawMline l, WALLCOLORS + lightlev
                    End If
                ElseIf sectors(back).floorheight <> sectors(front).floorheight Then
                    AM_drawMline l, FDWALLCOLORS + lightlev ' floor level change
                ElseIf sectors(back).ceilingheight <> sectors(front).ceilingheight Then
                    AM_drawMline l, CDWALLCOLORS + lightlev ' ceiling level change
                ElseIf cheating <> 0 Then
                    AM_drawMline l, TSWALLCOLORS + lightlev
                End If
            End If
        ElseIf players(plr).powers(pw_allmap) <> 0 Then
            If (flags And LINE_NEVERSEE) = 0 Then AM_drawMline l, GRAYS + 3
        End If
nextline:
    Next
End Sub

' Rotation in 2D.  Used to rotate player arrow line character.
Private Sub AM_rotate(x As Long, y As Long, ByVal a As Long)
    Dim tmpx As Long, fa As Long

    fa = FINEANGLE(a)
    tmpx = FixedMul(x, finecosine(fa)) - FixedMul(y, finesine(fa))

    y = FixedMul(x, finesine(fa)) + FixedMul(y, finecosine(fa))

    x = tmpx
End Sub

Private Sub AM_drawLineCharacter(lineguy() As mline_t, ByVal lineguylines As Long, ByVal scl As Long, _
                                 ByVal angle As Long, ByVal color As Long, ByVal x As Long, ByVal y As Long)
    Dim i As Long, l As mline_t

    For i = 0 To lineguylines - 1
        l.a.x = lineguy(i).a.x
        l.a.y = lineguy(i).a.y

        If scl <> 0 Then
            l.a.x = FixedMul(scl, l.a.x)
            l.a.y = FixedMul(scl, l.a.y)
        End If

        If angle <> 0 Then AM_rotate l.a.x, l.a.y, angle

        l.a.x = l.a.x + x
        l.a.y = l.a.y + y

        l.b.x = lineguy(i).b.x
        l.b.y = lineguy(i).b.y

        If scl <> 0 Then
            l.b.x = FixedMul(scl, l.b.x)
            l.b.y = FixedMul(scl, l.b.y)
        End If

        If angle <> 0 Then AM_rotate l.b.x, l.b.y, angle

        l.b.x = l.b.x + x
        l.b.y = l.b.y + y

        AM_drawMline l, color
    Next
End Sub

Private Sub AM_drawPlayers()
    Dim i As Long, their_color As Long, color As Long, mo As Long
    Dim their_colors(0 To 3) As Long

    If Not netgame Then
        mo = players(plr).mo
        If cheating <> 0 Then
            AM_drawLineCharacter cheat_player_arrow, 16, 0, mobjs(mo).angle, WHITE, mobjs(mo).x, mobjs(mo).y
        Else
            AM_drawLineCharacter player_arrow, 7, 0, mobjs(mo).angle, WHITE, mobjs(mo).x, mobjs(mo).y
        End If
        Exit Sub
    End If

    their_colors(0) = GREENS
    their_colors(1) = GRAYS
    their_colors(2) = BROWNS
    their_colors(3) = REDS
    their_color = -1

    For i = 0 To MAXPLAYERS - 1
        their_color = their_color + 1

        If (deathmatch <> 0 And Not singledemo) And i <> plr Then GoTo nextplayer

        If Not playeringame(i) Then GoTo nextplayer

        If players(i).powers(pw_invisibility) <> 0 Then
            color = 246                          ' *close* to black
        Else
            color = their_colors(their_color)
        End If

        mo = players(i).mo
        AM_drawLineCharacter player_arrow, 7, 0, mobjs(mo).angle, color, mobjs(mo).x, mobjs(mo).y
nextplayer:
    Next
End Sub

Private Sub AM_drawThings(ByVal colors As Long, ByVal colorrange As Long)
    Dim i As Long, t As Long

    For i = 0 To numsectors - 1
        t = sectors(i).thinglist
        Do While t <> 0
            AM_drawLineCharacter thintriangle_guy, 3, 16 * FRACUNIT, mobjs(t).angle, colors + lightlev, _
                                 mobjs(t).x, mobjs(t).y
            t = mobjs(t).snext
        Loop
    Next
End Sub

Private Sub AM_drawMarks()
    Dim i As Long, fx As Long, fy As Long, w As Long, h As Long

    For i = 0 To AM_NUMMARKPOINTS - 1
        If markpoints(i).x <> -1 Then
            '      w = SHORT(marknums[i]->width);
            '      h = SHORT(marknums[i]->height);
            w = 5                                ' because something's wrong with the wad, i guess
            h = 6                                ' because something's wrong with the wad, i guess
            fx = CXMTOF(markpoints(i).x)
            fy = CYMTOF(markpoints(i).y)
            If fx >= f_x And fx <= f_w - w And fy >= f_y And fy <= f_h - h Then
                V_DrawPatch fx, fy, FB, marknums(i)
            End If
        End If
    Next
End Sub

Private Sub AM_drawCrosshair(ByVal color As Long)
    screens((f_w * (f_h + 1)) \ 2) = color       ' single point for now
End Sub

Public Sub AM_Drawer()
    If Not automapactive Then Exit Sub

    AM_clearFB BACKGROUND
    If grid Then AM_drawGrid GRIDCOLORS
    AM_drawWalls
    AM_drawPlayers
    If cheating = 2 Then AM_drawThings THINGCOLORS, THINGRANGE
    AM_drawCrosshair XHAIRCOLORS

    AM_drawMarks

    V_MarkRect f_x, f_y, f_w, f_h
End Sub
