Attribute VB_Name = "i_video"
' i_video.bas - Win32 video and input (replaces the X11 i_video.c).
'
' The 320x200 8-bit frame buffer screens(0..63999) is copied into a palettized
' DIB section and stretched onto frmDoom.  Keyboard input arrives through a
' subclassed window procedure (so Alt / F10 do not activate the system menu)
' and is posted as Doom events.  The mouse is read relative to the window
' centre while the game is being played.  With smooth frames D_DoomLoop
' draws one frame per screen refresh, timed by the desktop compositor's
' refresh grid (I_WaitForRefresh), so frames reach the monitor evenly.
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO256
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As Long
End Type

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO256, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function AdjustWindowRect Lib "user32" (lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function ValidateRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GdiFlush Lib "gdi32" () As Long
Private Declare Function DwmGetCompositionTimingInfo Lib "dwmapi" (ByVal hWnd As Long, pTimingInfo As Any) As Long

Private Const GWL_WNDPROC As Long = -4
Private Const GWL_STYLE As Long = -16
Private Const WM_CLOSE As Long = &H10
Private Const WM_ACTIVATEAPP As Long = &H1C
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_SYSKEYDOWN As Long = &H104
Private Const WM_SYSKEYUP As Long = &H105
Private Const WM_SYSCHAR As Long = &H106
Private Const WM_SYSCOMMAND As Long = &H112
Private Const SC_KEYMENU As Long = &HF100&
Private Const SWP_NOMOVE As Long = 2
Private Const SWP_NOZORDER As Long = 4
Private Const COLORONCOLOR As Long = 3
Private Const SRCCOPY As Long = &HCC0020
Private Const VREFRESH As Long = 116
Private Const DWM_TIMING_SIZE As Long = 292      ' sizeof(DWM_TIMING_INFO), packed

Private hwndMain As Long
Private oldWndProc As Long
Private memDC As Long
Private hDib As Long
Private hOldBmp As Long
Private dibBits As Long
Private bmi As BITMAPINFO256
Private palQuads(0 To 255) As Long
Private keyIsDown(0 To 255) As Boolean
Private graphicsUp As Boolean

Private mouseGrabbed As Boolean
Private mouseCursorHidden As Boolean
Private lastButtons As Long
Private mouseEnabled As Boolean

Private refreshms As Double                      ' one screen refresh in ms (GetDeviceCaps)
Private framems As Double                        ' the moment the frame being drawn shows
Private framegrid As Double                      ' the refresh it follows (framems - offset)
Private nodwm As Boolean                         ' no dwmapi.dll (before Windows Vista)
Private dwmtiming(0 To DWM_TIMING_SIZE - 1) As Byte
Private framesshown As Long                      ' frames drawn (logged at shutdown)
Private firstframems As Double
Private clockpaced As Long                       ' refreshes timed by the clock alone
Private lateframes As Long                       ' frames that missed their refresh

Public currentPalette(0 To 767) As Byte          ' last palette passed to I_SetPalette (for screenshots)
Public framelog As Integer                       ' -framelog <file>: frame and tic timing
Private blitms As Double                         ' (for the frame log)
Public windowHidden As Boolean

Public Sub I_InitGraphics()
    Dim w As Long, h As Long, p As Long, rate As Long
    Dim r As RECT, hdc As Long

    If graphicsUp Then Exit Sub

    ' window size: -width N (client width, height = width * 3 / 4) or -scale N
    w = 640
    p = M_CheckParm("-scale")
    If p > 0 And p < myargc - 1 Then w = 320 * Val(myargv(p + 1))
    p = M_CheckParm("-width")
    If p > 0 And p < myargc - 1 Then w = Val(myargv(p + 1))
    If w < 320 Then w = 320
    h = (w * 3) \ 4

    windowHidden = (M_CheckParm("-hidden") > 0)
    mouseEnabled = (usemouse <> 0) And (M_CheckParm("-nomouse") = 0) And Not windowHidden

    Load frmDoom
    hwndMain = frmDoom.hWnd

    r.Left = 0: r.Top = 0: r.Right = w: r.Bottom = h
    AdjustWindowRect r, GetWindowLong(hwndMain, GWL_STYLE), 0
    SetWindowPos hwndMain, 0, (GetSystemMetrics(0) - (r.Right - r.Left)) \ 2, _
        (GetSystemMetrics(1) - (r.Bottom - r.Top)) \ 2, r.Right - r.Left, r.Bottom - r.Top, SWP_NOZORDER

    ' palettized 320x200 top-down DIB section
    With bmi.bmiHeader
        .biSize = Len(bmi.bmiHeader)
        .biWidth = SCREENWIDTH
        .biHeight = -SCREENHEIGHT
        .biPlanes = 1
        .biBitCount = 8
        .biCompression = 0
        .biClrUsed = 256
        .biClrImportant = 256
    End With
    hdc = GetDC(hwndMain)
    memDC = CreateCompatibleDC(hdc)
    hDib = CreateDIBSection(memDC, bmi, 0, dibBits, 0, 0)
    rate = GetDeviceCaps(hdc, VREFRESH)          ' 0 or 1 = the hardware default
    ReleaseDC hwndMain, hdc
    If rate < 30 Or rate > 500 Then rate = 60
    refreshms = 1000# / rate
    I_Printf "I_InitGraphics: " & w & "x" & h & " window, display " & rate & " Hz"
    If hDib = 0 Or dibBits = 0 Then I_Error "I_InitGraphics: CreateDIBSection failed"
    hOldBmp = SelectObject(memDC, hDib)

    oldWndProc = SetWindowLong(hwndMain, GWL_WNDPROC, AddressOf DoomWndProc)
    graphicsUp = True

    ' as the DOS version does: start with the normal palette
    I_SetPalette W_CacheLumpName("PLAYPAL")

    If Not windowHidden Then
        frmDoom.Show
        frmDoom.SetFocus
    End If
    DoEvents
End Sub

' Window title for a status such as waiting for net game players; "" puts
' the normal title back.
Public Sub I_SetWindowTitle(ByVal s As String)
    If Not graphicsUp Then Exit Sub
    If Len(s) = 0 Then s = "DoomVB"
    frmDoom.Caption = s
    DoEvents
End Sub

Public Sub I_ShutdownGraphics()
    Dim msg As String
    On Error Resume Next
    If Not graphicsUp Then Exit Sub
    graphicsUp = False
    If framesshown > 1 Then
        msg = "I_ShutdownGraphics: " & framesshown & " frames, " & _
              Format$((framesshown - 1) * 1000# / (I_GetTimeExact() - firstframems), "0.0") & " fps"
        If smoothactive Then msg = msg & " (" & lateframes & " late, " & clockpaced & " timed by the clock)"
        I_Printf msg
    End If
    ReleaseMouse
    If oldWndProc <> 0 Then SetWindowLong hwndMain, GWL_WNDPROC, oldWndProc
    oldWndProc = 0
    If memDC <> 0 Then
        SelectObject memDC, hOldBmp
        DeleteObject hDib
        DeleteDC memDC
    End If
    memDC = 0
    Unload frmDoom
End Sub

' Takes the offset of a 768 byte palette inside wadbuf() (C: byte* palette).
Public Sub I_SetPalette(ByVal palofs As Long)
    Dim i As Long, r As Long, g As Long, b As Long
    CopyMemory currentPalette(0), wadbuf(palofs), 768
    For i = 0 To 255
        r = gammatable(usegamma, currentPalette(i * 3))
        g = gammatable(usegamma, currentPalette(i * 3 + 1))
        b = gammatable(usegamma, currentPalette(i * 3 + 2))
        palQuads(i) = b Or (g * 256&) Or (r * 65536)
    Next
    If memDC <> 0 Then SetDIBColorTable memDC, 0, 256, palQuads(0)
End Sub

Public Sub I_UpdateNoBlit()
End Sub

Public Sub I_StartFrame()
End Sub

' Copies screens(0) to the window.
Public Sub I_FinishUpdate()
    Dim hdc As Long, r As RECT
    If Not graphicsUp Then Exit Sub
    CopyMemory ByVal dibBits, screens(0), SCREENSIZE
    If framesshown = 0 Then firstframems = I_GetTimeExact()
    framesshown = framesshown + 1
    If Not windowHidden Then BlitToWindow
    If framelog <> 0 Then blitms = I_GetTimeExact()
End Sub

' Smooth frames (D_DoomLoop, after every frame): waits until shortly after
' the next screen refresh.  The frame drawn then reaches the compositor well
' before its following pass, and it shows the game as of that moment
' (I_FrameTime), so every frame moves the picture by the same amount.  The
' wait keeps building tics on time (the mouse is read when a tic is built)
' and feeds the sound, as TryRunTics' own wait loop does.
Public Sub I_WaitForRefresh()
    Dim nowms As Double, grid As Double, period As Double, offset As Double, target As Double

    nowms = I_GetTimeExact()
    If Not RefreshGrid(grid, period) Then
        clockpaced = clockpaced + 1              ' no compositor: keep the last frame's rhythm
        grid = framegrid
        period = refreshms
    End If
    offset = period / 4
    If offset > 3 Then offset = 3
    target = grid + offset + period * (Int((nowms - grid - offset) / period) + 1)
    framegrid = target - offset
    If framems > 0 And target - framems > period * 1.5 Then lateframes = lateframes + 1
    If framelog <> 0 Then
        Print #framelog, "F " & gametic & " " & Format$(framems, "0.000") & " " & Format$(blitms, "0.000") & _
                         " " & Format$(target, "0.000") & " " & Format$(period, "0.000")
    End If

    Do
        NetUpdate                                ' build tics on time
        I_UpdateSound
        nowms = I_GetTimeExact()
        If nowms >= target Then Exit Do
        If target - nowms > 1.5 Then Sleep 1 Else Sleep 0
    Loop
    framems = target
End Sub

' Smooth frames: the moment the frame being drawn shows the game at
' (I_GetTimeExact milliseconds).
Public Function I_FrameTime() As Double
    If framems = 0 Then framems = I_GetTimeExact()
    I_FrameTime = framems
End Function

' The compositor's refresh grid: the time of one vertical blank and the
' refresh period, in I_GetTimeExact milliseconds.  False without one to ask
' (-hidden, no dwmapi.dll, composition off, no performance counter).
Private Function RefreshGrid(grid As Double, period As Double) As Boolean
    Dim sz As Long, vblank As Currency, qpcperiod As Currency
    If windowHidden Or nodwm Then Exit Function
    On Error GoTo missing
    sz = DWM_TIMING_SIZE
    CopyMemory dwmtiming(0), sz, 4               ' cbSize
    If DwmGetCompositionTimingInfo(0, dwmtiming(0)) <> 0 Then Exit Function
    CopyMemory qpcperiod, dwmtiming(12), 8       ' qpcRefreshPeriod
    CopyMemory vblank, dwmtiming(28), 8          ' qpcVBlank
    grid = I_CountToMS(vblank)
    If grid < 0 Then Exit Function
    period = I_CountToMS(vblank + qpcperiod) - grid
    If period < 2 Or period > 50 Then Exit Function
    RefreshGrid = True
    Exit Function
missing:
    nodwm = True
End Function

Private Sub BlitToWindow()
    Dim hdc As Long, r As RECT
    GetClientRect hwndMain, r
    hdc = GetDC(hwndMain)
    SetStretchBltMode hdc, COLORONCOLOR
    StretchBlt hdc, 0, 0, r.Right, r.Bottom, memDC, 0, 0, SCREENWIDTH, SCREENHEIGHT, SRCCOPY
    GdiFlush                                     ' hand it over now, not at the next GDI call
    ReleaseDC hwndMain, hdc
End Sub

' Copies screens(0) into screen number scrn (C: I_ReadScreen(byte* scr)).
Public Sub I_ReadScreen(ByVal scrn As Long)
    CopyMemory screens(scrn * SCREENSIZE), screens(0), SCREENSIZE
End Sub

' Called once per tic: pumps window messages (keyboard events are posted by
' DoomWndProc) and reads the mouse.
Public Sub I_StartTic()
    DoEvents
    If quitRequested Then I_Quit
    If mouseEnabled Then PollMouse
End Sub

Private Function WantMouse() As Boolean
    WantMouse = False
    If GetForegroundWindow() <> hwndMain Then Exit Function
    If menuactive Or paused Or demoplayback Then Exit Function
    If gamestate <> GS_LEVEL Then Exit Function
    WantMouse = True
End Function

Private Sub GrabMouse()
    Dim r As RECT, p As POINTAPI
    GetClientRect hwndMain, r
    p.x = 0: p.y = 0
    ClientToScreen hwndMain, p
    r.Left = r.Left + p.x: r.Right = r.Right + p.x
    r.Top = r.Top + p.y: r.Bottom = r.Bottom + p.y
    ClipCursor r
    If Not mouseCursorHidden Then
        ShowCursor 0
        mouseCursorHidden = True
    End If
    CenterMouse
    mouseGrabbed = True
End Sub

Private Sub ReleaseMouse()
    ClipCursor ByVal 0&
    If mouseCursorHidden Then
        ShowCursor 1
        mouseCursorHidden = False
    End If
    mouseGrabbed = False
End Sub

Private Sub WindowCenter(p As POINTAPI)
    Dim r As RECT
    GetClientRect hwndMain, r
    p.x = r.Right \ 2
    p.y = r.Bottom \ 2
    ClientToScreen hwndMain, p
End Sub

Private Sub CenterMouse()
    Dim c As POINTAPI
    WindowCenter c
    SetCursorPos c.x, c.y
End Sub

Private Sub PollMouse()
    Dim p As POINTAPI, c As POINTAPI, ev As event_t, buttons As Long
    If Not WantMouse() Then
        If mouseGrabbed Then ReleaseMouse
        Exit Sub
    End If
    If Not mouseGrabbed Then
        GrabMouse
        Exit Sub
    End If
    GetCursorPos p
    WindowCenter c
    buttons = 0
    If GetAsyncKeyState(1) < 0 Then buttons = buttons Or 1     ' VK_LBUTTON
    If GetAsyncKeyState(2) < 0 Then buttons = buttons Or 2     ' VK_RBUTTON
    If GetAsyncKeyState(4) < 0 Then buttons = buttons Or 4     ' VK_MBUTTON
    If p.x <> c.x Or p.y <> c.y Or buttons <> lastButtons Then
        ev.typ = ev_mouse
        ev.data1 = buttons
        ev.data2 = (p.x - c.x) * 4
        ev.data3 = (c.y - p.y) * 4
        D_PostEvent ev
        lastButtons = buttons
        SetCursorPos c.x, c.y
    End If
End Sub

' Windows virtual key -> Doom key code (0 = ignore)
Private Function TranslateKey(ByVal vk As Long) As Long
    Select Case vk
        Case 37: TranslateKey = KEY_LEFTARROW
        Case 38: TranslateKey = KEY_UPARROW
        Case 39: TranslateKey = KEY_RIGHTARROW
        Case 40: TranslateKey = KEY_DOWNARROW
        Case 27: TranslateKey = KEY_ESCAPE
        Case 13: TranslateKey = KEY_ENTER
        Case 9: TranslateKey = KEY_TAB
        Case 8: TranslateKey = KEY_BACKSPACE
        Case 19: TranslateKey = KEY_PAUSE
        Case 112 To 121: TranslateKey = KEY_F1 + (vk - 112)
        Case 122: TranslateKey = KEY_F11
        Case 123: TranslateKey = KEY_F12
        Case 16: TranslateKey = KEY_RSHIFT
        Case 17: TranslateKey = KEY_RCTRL
        Case 18: TranslateKey = KEY_RALT
        Case 32: TranslateKey = 32
        Case 48 To 57: TranslateKey = vk
        Case 65 To 90: TranslateKey = vk + 32
        Case 96 To 105: TranslateKey = 48 + (vk - 96)              ' numpad digits
        Case 187, 107: TranslateKey = KEY_EQUALS
        Case 189, 109: TranslateKey = KEY_MINUS
        Case 188: TranslateKey = 44                                 ' ,
        Case 190: TranslateKey = 46                                 ' .
        Case 191: TranslateKey = 47                                 ' /
        Case 186: TranslateKey = 59                                 ' ;
        Case 222: TranslateKey = 39                                 ' '
        Case 219: TranslateKey = 91                                 ' [
        Case 221: TranslateKey = 93                                 ' ]
        Case 220: TranslateKey = 92                                 ' \
        Case 192: TranslateKey = 96                                 ' `
        Case Else: TranslateKey = 0
    End Select
End Function

Private Sub ReleaseAllKeys()
    Dim i As Long, ev As event_t
    For i = 0 To 255
        If keyIsDown(i) Then
            keyIsDown(i) = False
            ev.typ = ev_keyup
            ev.data1 = i
            D_PostEvent ev
        End If
    Next
End Sub

Public Function DoomWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim k As Long, ev As event_t
    Select Case uMsg
        Case WM_KEYDOWN, WM_SYSKEYDOWN
            ' let Alt+F4 close the window
            If wParam = 115 And uMsg = WM_SYSKEYDOWN Then
                quitRequested = True
                DoomWndProc = 0
                Exit Function
            End If
            ' a -hidden run takes its input from -key only: the invisible
            ' window can still own the keyboard focus
            If windowHidden Then k = 0 Else k = TranslateKey(wParam)
            If k <> 0 Then
                keyIsDown(k) = True
                ev.typ = ev_keydown
                ev.data1 = k
                D_PostEvent ev
            End If
            DoomWndProc = 0
            Exit Function
        Case WM_KEYUP, WM_SYSKEYUP
            If windowHidden Then k = 0 Else k = TranslateKey(wParam)
            If k <> 0 Then
                keyIsDown(k) = False
                ev.typ = ev_keyup
                ev.data1 = k
                D_PostEvent ev
            End If
            DoomWndProc = 0
            Exit Function
        Case WM_CHAR, WM_SYSCHAR
            DoomWndProc = 0
            Exit Function
        Case WM_SYSCOMMAND
            If (wParam And &HFFF0&) = SC_KEYMENU Then
                DoomWndProc = 0
                Exit Function
            End If
        Case WM_KILLFOCUS
            ReleaseAllKeys
            If mouseGrabbed Then ReleaseMouse
        Case WM_ACTIVATEAPP
            If wParam = 0 Then
                ReleaseAllKeys
                If mouseGrabbed Then ReleaseMouse
            End If
        Case WM_ERASEBKGND
            DoomWndProc = 1
            Exit Function
        Case WM_CLOSE
            quitRequested = True
            DoomWndProc = 0
            Exit Function
    End Select
    DoomWndProc = CallWindowProc(oldWndProc, hWnd, uMsg, wParam, lParam)
End Function

' called from frmDoom_Paint
Public Sub I_RepaintWindow()
    If graphicsUp And Not windowHidden Then BlitToWindow
End Sub
