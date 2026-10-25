Attribute VB_Name = "hu_lib"
' hu_lib.bas - heads-up text and input code (hu_lib.c).
'
' The widgets always use hu_font (C: patch_t** f).  The C widgets kept a
' pointer to the boolean that switches them on; here the draw and erase
' routines take that boolean's current value instead (C: boolean* on).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const FG As Long = 0                     ' screen the widgets draw on

Public Const HU_CHARERASE As Long = KEY_BACKSPACE
Public Const HU_MAXLINES As Long = 4
Public Const HU_MAXLINELENGTH As Long = 80

' text line widget (parent of scrolling text and input text widgets)
Public Type hu_textline_t
    x As Long                                    ' left-justified position of scrolling text window
    y As Long
    sc As Long                                   ' start character
    txt As String                                ' line of text (C: l[] and len)
    needsupdate As Long                          ' whether this line needs to be updated
End Type

' scrolling text window widget (child of text line widget)
Public Type hu_stext_t
    l(0 To 3) As hu_textline_t                   ' text lines to draw (HU_MAXLINES)
    h As Long                                    ' height in lines
    cl As Long                                   ' current line number
    laston As Boolean                            ' last value of on
End Type

' input text line widget (child of text line widget)
Public Type hu_itext_t
    l As hu_textline_t                           ' text line to input on
    lm As Long                                   ' left margin past which I am not to delete characters
    laston As Boolean                            ' last value of on
End Type

Public Sub HUlib_init()
End Sub

Public Sub HUlib_clearTextLine(t As hu_textline_t)
    t.txt = ""
    t.needsupdate = 1                            ' true
End Sub

Public Sub HUlib_initTextLine(t As hu_textline_t, ByVal x As Long, ByVal y As Long, ByVal sc As Long)
    t.x = x
    t.y = y
    t.sc = sc
    HUlib_clearTextLine t
End Sub

Public Function HUlib_addCharToTextLine(t As hu_textline_t, ByVal ch As Long) As Boolean
    If Len(t.txt) = HU_MAXLINELENGTH Then
        HUlib_addCharToTextLine = False
    Else
        t.txt = t.txt & Chr$(ch)
        t.needsupdate = 4
        HUlib_addCharToTextLine = True
    End If
End Function

Public Function HUlib_delCharFromTextLine(t As hu_textline_t) As Boolean
    If Len(t.txt) = 0 Then
        HUlib_delCharFromTextLine = False
    Else
        t.txt = Left$(t.txt, Len(t.txt) - 1)
        t.needsupdate = 4
        HUlib_delCharFromTextLine = True
    End If
End Function

Public Sub HUlib_drawTextLine(l As hu_textline_t, ByVal drawcursor As Boolean)
    Dim i As Long, w As Long, x As Long, c As Long

    ' draw the new stuff
    x = l.x
    For i = 1 To Len(l.txt)
        c = toupper(Asc(Mid$(l.txt, i, 1)))
        If c <> 32 And c >= l.sc And c <= 95 Then  ' '_'
            w = PatchWidth(hu_font(c - l.sc))
            If x + w > SCREENWIDTH Then Exit For
            V_DrawPatchDirect x, l.y, FG, hu_font(c - l.sc)
            x = x + w
        Else
            x = x + 4
            If x >= SCREENWIDTH Then Exit For
        End If
    Next

    ' draw the cursor if requested
    If drawcursor Then
        If x + PatchWidth(hu_font(95 - l.sc)) <= SCREENWIDTH Then
            V_DrawPatchDirect x, l.y, FG, hu_font(95 - l.sc)
        End If
    End If
End Sub

' sorta called by HU_Erase and just better darn get things straight
Public Sub HUlib_eraseTextLine(l As hu_textline_t)
    Dim lh As Long, y As Long, yoffset As Long

    ' Only erases when NOT in automap and the screen is reduced,
    ' and the text must either need updating or refreshing
    ' (because of a recent change back from the automap)
    If Not automapactive And viewwindowx <> 0 And l.needsupdate <> 0 Then
        lh = PatchHeight(hu_font(0)) + 1
        yoffset = l.y * SCREENWIDTH
        For y = l.y To l.y + lh - 1
            If y < viewwindowy Or y >= viewwindowy + viewheight Then
                R_VideoErase yoffset, SCREENWIDTH                          ' erase entire line
            Else
                R_VideoErase yoffset, viewwindowx                          ' erase left border
                R_VideoErase yoffset + viewwindowx + viewwidth, viewwindowx ' erase right border
            End If
            yoffset = yoffset + SCREENWIDTH
        Next
    End If

    If l.needsupdate <> 0 Then l.needsupdate = l.needsupdate - 1
End Sub

Public Sub HUlib_initSText(s As hu_stext_t, ByVal x As Long, ByVal y As Long, ByVal h As Long, ByVal startchar As Long)
    Dim i As Long
    s.h = h
    s.laston = True
    s.cl = 0
    For i = 0 To h - 1
        HUlib_initTextLine s.l(i), x, y - i * (PatchHeight(hu_font(0)) + 1), startchar
    Next
End Sub

Public Sub HUlib_addLineToSText(s As hu_stext_t)
    Dim i As Long

    ' add a clear line
    s.cl = s.cl + 1
    If s.cl = s.h Then s.cl = 0
    HUlib_clearTextLine s.l(s.cl)

    ' everything needs updating
    For i = 0 To s.h - 1
        s.l(i).needsupdate = 4
    Next
End Sub

Public Sub HUlib_addMessageToSText(s As hu_stext_t, ByVal prefix As String, ByVal msg As String)
    Dim i As Long
    HUlib_addLineToSText s
    For i = 1 To Len(prefix)
        HUlib_addCharToTextLine s.l(s.cl), Asc(Mid$(prefix, i, 1))
    Next
    For i = 1 To Len(msg)
        HUlib_addCharToTextLine s.l(s.cl), Asc(Mid$(msg, i, 1))
    Next
End Sub

Public Sub HUlib_drawSText(s As hu_stext_t, ByVal isOn As Boolean)
    Dim i As Long, idx As Long

    If Not isOn Then Exit Sub                    ' if not on, don't draw

    ' draw everything
    For i = 0 To s.h - 1
        idx = s.cl - i
        If idx < 0 Then idx = idx + s.h          ' handle queue of lines

        ' need a decision made here on whether to skip the draw
        HUlib_drawTextLine s.l(idx), False       ' no cursor, please
    Next
End Sub

Public Sub HUlib_eraseSText(s As hu_stext_t, ByVal isOn As Boolean)
    Dim i As Long
    For i = 0 To s.h - 1
        If s.laston And Not isOn Then s.l(i).needsupdate = 4
        HUlib_eraseTextLine s.l(i)
    Next
    s.laston = isOn
End Sub

Public Sub HUlib_initIText(it As hu_itext_t, ByVal x As Long, ByVal y As Long, ByVal startchar As Long)
    it.lm = 0                                    ' default left margin is start of text
    it.laston = True
    HUlib_initTextLine it.l, x, y, startchar
End Sub

' The following deletion routines adhere to the left margin restriction
Public Sub HUlib_delCharFromIText(it As hu_itext_t)
    If Len(it.l.txt) <> it.lm Then HUlib_delCharFromTextLine it.l
End Sub

Public Sub HUlib_eraseLineFromIText(it As hu_itext_t)
    Do While it.lm <> Len(it.l.txt)
        HUlib_delCharFromTextLine it.l
    Loop
End Sub

' Resets left margin as well
Public Sub HUlib_resetIText(it As hu_itext_t)
    it.lm = 0
    HUlib_clearTextLine it.l
End Sub

Public Sub HUlib_addPrefixToIText(it As hu_itext_t, ByVal prefix As String)
    Dim i As Long
    For i = 1 To Len(prefix)
        HUlib_addCharToTextLine it.l, Asc(Mid$(prefix, i, 1))
    Next
    it.lm = Len(it.l.txt)
End Sub

' wrapper function for handling general keyed input.
' returns true if it ate the key
Public Function HUlib_keyInIText(it As hu_itext_t, ByVal ch As Long) As Boolean
    If ch >= 32 And ch <= 95 Then                ' ' ' .. '_'
        HUlib_addCharToTextLine it.l, ch
    ElseIf ch = KEY_BACKSPACE Then
        HUlib_delCharFromIText it
    ElseIf ch <> KEY_ENTER Then
        HUlib_keyInIText = False                 ' did not eat key
        Exit Function
    End If
    HUlib_keyInIText = True                      ' ate the key
End Function

Public Sub HUlib_drawIText(it As hu_itext_t, ByVal isOn As Boolean)
    If Not isOn Then Exit Sub
    HUlib_drawTextLine it.l, True                ' draw the line w/ cursor
End Sub

Public Sub HUlib_eraseIText(it As hu_itext_t, ByVal isOn As Boolean)
    If it.laston And Not isOn Then it.l.needsupdate = 4
    HUlib_eraseTextLine it.l
    it.laston = isOn
End Sub
