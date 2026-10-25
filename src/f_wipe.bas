Attribute VB_Name = "f_wipe"
' f_wipe.bas - mission begin melt/wipe screen special effect (f_wipe.c).
'
' The start screen is kept in screens 2, the end screen in screen 3 and the
' wipe is drawn on screen 0, as in the C version.  The melt moves 16-bit
' pixel pairs ("shorts"); short k of a screen is bytes 2k and 2k+1.
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' SCREEN WIPE PACKAGE

' when zero, stop the wipe
Private go As Boolean

' byte offsets of the screens in screens()
Private wipe_scr_start As Long
Private wipe_scr_end As Long
Private wipe_scr As Long

' column positions of the melt (C: static int* y)
Private ypos() As Long

' screen at byte offset ofs: short array of width x height, row major ->
' column major
Private Sub wipe_shittyColMajorXform(ByVal ofs As Long, ByVal width As Long, ByVal height As Long)
    Dim x As Long, y As Long, dest() As Byte, s As Long, d As Long

    ReDim dest(0 To width * height * 2 - 1)
    For y = 0 To height - 1
        For x = 0 To width - 1
            s = ofs + (y * width + x) * 2
            d = (x * height + y) * 2
            dest(d) = screens(s)
            dest(d + 1) = screens(s + 1)
        Next
    Next
    CopyMemory screens(ofs), dest(0), width * height * 2
End Sub

Private Function wipe_initColorXForm(ByVal width As Long, ByVal height As Long, ByVal ticks As Long) As Boolean
    CopyMemory screens(wipe_scr), screens(wipe_scr_start), width * height
    wipe_initColorXForm = False
End Function

Private Function wipe_doColorXForm(ByVal width As Long, ByVal height As Long, ByVal ticks As Long) As Boolean
    Dim changed As Boolean, i As Long, w As Long, e As Long, newval As Long

    changed = False

    For i = 0 To width * height - 1
        w = screens(wipe_scr + i)
        e = screens(wipe_scr_end + i)
        If w <> e Then
            If w > e Then
                newval = w - ticks
                If newval < e Then
                    screens(wipe_scr + i) = e
                Else
                    screens(wipe_scr + i) = newval
                End If
                changed = True
            ElseIf w < e Then
                newval = w + ticks
                If newval > e Then
                    screens(wipe_scr + i) = e
                Else
                    screens(wipe_scr + i) = newval
                End If
                changed = True
            End If
        End If
    Next
    wipe_doColorXForm = Not changed
End Function

Private Function wipe_exitColorXForm(ByVal width As Long, ByVal height As Long, ByVal ticks As Long) As Boolean
    wipe_exitColorXForm = False
End Function

Private Function wipe_initMelt(ByVal width As Long, ByVal height As Long, ByVal ticks As Long) As Boolean
    Dim i As Long, r As Long

    ' copy start screen to main screen
    CopyMemory screens(wipe_scr), screens(wipe_scr_start), width * height

    ' makes this wipe faster (in theory)
    ' to have stuff in column-major format
    wipe_shittyColMajorXform wipe_scr_start, width \ 2, height
    wipe_shittyColMajorXform wipe_scr_end, width \ 2, height

    ' setup initial column positions
    ' (y<0 => not ready to scroll yet)
    ReDim ypos(0 To width - 1)
    ypos(0) = -(M_Random() Mod 16)
    For i = 1 To width - 1
        r = (M_Random() Mod 3) - 1
        ypos(i) = ypos(i - 1) + r
        If ypos(i) > 0 Then
            ypos(i) = 0
        ElseIf ypos(i) = -16 Then
            ypos(i) = -15
        End If
    Next
    wipe_initMelt = False
End Function

Private Function wipe_doMelt(ByVal width As Long, ByVal height As Long, ByVal ticks As Long) As Boolean
    Dim i As Long, j As Long, dy As Long, s As Long, d As Long, done As Boolean

    done = True
    width = width \ 2

    Do While ticks > 0
        ticks = ticks - 1
        For i = 0 To width - 1
            If ypos(i) < 0 Then
                ypos(i) = ypos(i) + 1
                done = False
            ElseIf ypos(i) < height Then
                If ypos(i) < 16 Then dy = ypos(i) + 1 Else dy = 8
                If ypos(i) + dy >= height Then dy = height - ypos(i)
                s = wipe_scr_end + (i * height + ypos(i)) * 2
                d = wipe_scr + (ypos(i) * width + i) * 2
                For j = dy To 1 Step -1
                    screens(d) = screens(s)
                    screens(d + 1) = screens(s + 1)
                    s = s + 2
                    d = d + width * 2
                Next
                ypos(i) = ypos(i) + dy
                s = wipe_scr_start + (i * height) * 2
                d = wipe_scr + (ypos(i) * width + i) * 2
                For j = height - ypos(i) To 1 Step -1
                    screens(d) = screens(s)
                    screens(d + 1) = screens(s + 1)
                    s = s + 2
                    d = d + width * 2
                Next
                done = False
            End If
        Next
    Loop

    wipe_doMelt = done
End Function

Private Function wipe_exitMelt(ByVal width As Long, ByVal height As Long, ByVal ticks As Long) As Boolean
    Erase ypos
    wipe_exitMelt = False
End Function

Public Function wipe_StartScreen(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long
    wipe_scr_start = 2 * SCREENSIZE
    I_ReadScreen 2
    wipe_StartScreen = 0
End Function

Public Function wipe_EndScreen(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long
    wipe_scr_end = 3 * SCREENSIZE
    I_ReadScreen 3
    V_DrawBlock x, y, 0, width, height, screens, wipe_scr_start ' restore start scr.
    wipe_EndScreen = 0
End Function

' (C: the wipes[] function pointer table; wipeno is wipe_ColorXForm or wipe_Melt)
Public Function wipe_ScreenWipe(ByVal wipeno As Long, ByVal x As Long, ByVal y As Long, _
                                ByVal width As Long, ByVal height As Long, ByVal ticks As Long) As Boolean
    Dim rc As Boolean

    ' initial stuff
    If Not go Then
        go = True
        wipe_scr = 0                             ' screens[0]
        If wipeno = wipe_Melt Then
            wipe_initMelt width, height, ticks
        Else
            wipe_initColorXForm width, height, ticks
        End If
    End If

    ' do a piece of wipe-in
    V_MarkRect 0, 0, width, height
    If wipeno = wipe_Melt Then
        rc = wipe_doMelt(width, height, ticks)
    Else
        rc = wipe_doColorXForm(width, height, ticks)
    End If

    ' final stuff
    If rc Then
        go = False
        If wipeno = wipe_Melt Then
            wipe_exitMelt width, height, ticks
        Else
            wipe_exitColorXForm width, height, ticks
        End If
    End If

    wipe_ScreenWipe = Not go
End Function
