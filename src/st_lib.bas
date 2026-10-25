Attribute VB_Name = "st_lib"
' st_lib.bas - the status bar widget code (st_lib.c).
'
' A widget keeps a copy of its patch list (lump numbers).  The C widgets also
' kept pointers to the value they show and to the boolean that switches them
' on; here the update routines take those values instead.
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' Background and foreground screen numbers
Private Const BG As Long = 4
Private Const FG As Long = 0

' st_stuff.h: location of the status bar
Public Const ST_HEIGHT As Long = 32              ' 32 * SCREEN_MUL
Public Const ST_WIDTH As Long = SCREENWIDTH
Public Const ST_Y As Long = SCREENHEIGHT - ST_HEIGHT

' Number widget
Public Type st_number_t
    x As Long                                    ' upper right-hand corner of the number (right-justified)
    y As Long
    numwidth As Long                             ' max # of digits in number (C: width)
    oldnum As Long                               ' last number value
    p(0 To 9) As Long                            ' list of patches for 0-9
    data As Long                                 ' user data
End Type

' Percent widget ("child" of number widget, or, more precisely,
' contains a number widget.)
Public Type st_percent_t
    n As st_number_t                             ' number information
    p As Long                                    ' percent sign graphic
End Type

' Multiple Icon widget
Public Type st_multicon_t
    x As Long                                    ' center-justified location of icons
    y As Long
    oldinum As Long                              ' last icon number
    p(0 To 41) As Long                           ' list of icons (up to ST_NUMFACES)
    data As Long                                 ' user data
End Type

' Binary Icon widget
Public Type st_binicon_t
    x As Long                                    ' center-justified location of icon
    y As Long
    oldval As Boolean                            ' last icon value
    p As Long                                    ' icon
    data As Long                                 ' user data
End Type

' Hack display negative frags.  Loads and store the stminus lump.
Private sttminus As Long

Public Sub STlib_init()
    sttminus = W_GetNumForName("STTMINUS")
End Sub

Public Sub STlib_initNum(n As st_number_t, ByVal x As Long, ByVal y As Long, pl() As Long, ByVal numwidth As Long)
    Dim i As Long
    n.x = x
    n.y = y
    n.oldnum = 0
    n.numwidth = numwidth
    For i = 0 To 9
        n.p(i) = pl(i)
    Next
End Sub

' A fairly efficient way to draw a number based on differences from the
' old number.  Note: worth the trouble?
Private Sub STlib_drawNum(n As st_number_t, ByVal num As Long, ByVal refresh As Boolean)
    Dim numdigits As Long, w As Long, h As Long, x As Long, neg As Boolean

    numdigits = n.numwidth
    w = PatchWidth(n.p(0))
    h = PatchHeight(n.p(0))

    n.oldnum = num

    neg = num < 0

    If neg Then
        If numdigits = 2 And num < -9 Then
            num = -9
        ElseIf numdigits = 3 And num < -99 Then
            num = -99
        End If
        num = -num
    End If

    ' clear the area
    x = n.x - numdigits * w

    If n.y - ST_Y < 0 Then I_Error "drawNum: n->y - ST_Y < 0"

    V_CopyRect x, n.y - ST_Y, BG, w * numdigits, h, x, n.y, FG

    ' if non-number, do not draw it
    If num = 1994 Then Exit Sub

    x = n.x

    ' in the special case of 0, you draw 0
    If num = 0 Then V_DrawPatch x - w, n.y, FG, n.p(0)

    ' draw the new number
    Do While num <> 0
        If numdigits = 0 Then Exit Do
        numdigits = numdigits - 1
        x = x - w
        V_DrawPatch x, n.y, FG, n.p(num Mod 10)
        num = num \ 10
    Loop

    ' draw a minus sign if necessary
    If neg Then V_DrawPatch x - 8, n.y, FG, sttminus
End Sub

Public Sub STlib_updateNum(n As st_number_t, ByVal num As Long, ByVal isOn As Boolean, ByVal refresh As Boolean)
    If isOn Then STlib_drawNum n, num, refresh
End Sub

Public Sub STlib_initPercent(p As st_percent_t, ByVal x As Long, ByVal y As Long, pl() As Long, ByVal percent As Long)
    STlib_initNum p.n, x, y, pl, 3
    p.p = percent
End Sub

Public Sub STlib_updatePercent(per As st_percent_t, ByVal num As Long, ByVal isOn As Boolean, ByVal refresh As Boolean)
    If refresh And isOn Then V_DrawPatch per.n.x, per.n.y, FG, per.p

    STlib_updateNum per.n, num, isOn, refresh
End Sub

' il(ilbase) .. il(ilbase + count - 1) is the icon list
Public Sub STlib_initMultIcon(mi As st_multicon_t, ByVal x As Long, ByVal y As Long, il() As Long, _
                              ByVal ilbase As Long, ByVal count As Long)
    Dim i As Long
    mi.x = x
    mi.y = y
    mi.oldinum = -1
    For i = 0 To count - 1
        mi.p(i) = il(ilbase + i)
    Next
End Sub

Public Sub STlib_updateMultIcon(mi As st_multicon_t, ByVal inum As Long, ByVal isOn As Boolean, ByVal refresh As Boolean)
    Dim w As Long, h As Long, x As Long, y As Long

    If isOn And (mi.oldinum <> inum Or refresh) And inum <> -1 Then
        If mi.oldinum <> -1 Then
            x = mi.x - PatchLeftOffset(mi.p(mi.oldinum))
            y = mi.y - PatchTopOffset(mi.p(mi.oldinum))
            w = PatchWidth(mi.p(mi.oldinum))
            h = PatchHeight(mi.p(mi.oldinum))

            If y - ST_Y < 0 Then I_Error "updateMultIcon: y - ST_Y < 0"

            V_CopyRect x, y - ST_Y, BG, w, h, x, y, FG
        End If
        V_DrawPatch mi.x, mi.y, FG, mi.p(inum)
        mi.oldinum = inum
    End If
End Sub

Public Sub STlib_initBinIcon(b As st_binicon_t, ByVal x As Long, ByVal y As Long, ByVal i As Long)
    b.x = x
    b.y = y
    b.oldval = False
    b.p = i
End Sub

Public Sub STlib_updateBinIcon(bi As st_binicon_t, ByVal val As Boolean, ByVal isOn As Boolean, ByVal refresh As Boolean)
    Dim x As Long, y As Long, w As Long, h As Long

    If isOn And (bi.oldval <> val Or refresh) Then
        x = bi.x - PatchLeftOffset(bi.p)
        y = bi.y - PatchTopOffset(bi.p)
        w = PatchWidth(bi.p)
        h = PatchHeight(bi.p)

        If y - ST_Y < 0 Then I_Error "updateBinIcon: y - ST_Y < 0"

        If val Then
            V_DrawPatch bi.x, bi.y, FG, bi.p
        Else
            V_CopyRect x, y - ST_Y, BG, w, h, x, y, FG
        End If

        bi.oldval = val
    End If
End Sub
