Attribute VB_Name = "m_misc"
' m_misc.bas - default.cfg handling, file helpers, screenshots (m_misc.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public Const SAVEGAMENAME As String = "doomsav"

Public numChannels As Long                       ' snd_channels (s_sound)
Public chat_macros(0 To 9) As String

Private Type BMPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Private Type BMPINFOHEADER
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

' ------------------------------------------------------------------ files
Public Function M_WriteFile(ByVal name As String, source() As Byte, ByVal length As Long) As Boolean
    Dim f As Integer
    On Error GoTo failed
    If Dir$(name) <> "" Then Kill name
    f = FreeFile
    Open name For Binary Access Write As #f
    If length > 0 Then
        Dim tmp() As Byte
        ReDim tmp(0 To length - 1)
        CopyMemory tmp(0), source(LBound(source)), length
        Put #f, 1, tmp
    End If
    Close #f
    M_WriteFile = True
    Exit Function
failed:
    M_WriteFile = False
End Function

' Reads a whole file into buffer() and returns its length (I_Error if missing).
Public Function M_ReadFile(ByVal name As String, buffer() As Byte) As Long
    Dim f As Integer, n As Long
    If Dir$(name) = "" Then I_Error "Couldn't read file " & name
    f = FreeFile
    Open name For Binary Access Read As #f
    n = LOF(f)
    If n > 0 Then
        ReDim buffer(0 To n - 1)
        Get #f, 1, buffer
    Else
        Erase buffer
    End If
    Close #f
    M_ReadFile = n
End Function

' ------------------------------------------------------------------ defaults
' name, pointer (by name), default value
Private Sub SetDefault(ByVal name As String, ByVal v As String)
    Select Case LCase$(name)
        Case "mouse_sensitivity": mouseSensitivity = Val(v)
        Case "sfx_volume": snd_SfxVolume = Val(v)
        Case "music_volume": snd_MusicVolume = Val(v)
        Case "show_messages": showMessages = Val(v)
        Case "key_right": key_right = Val(v)
        Case "key_left": key_left = Val(v)
        Case "key_up": key_up = Val(v)
        Case "key_down": key_down = Val(v)
        Case "key_strafeleft": key_strafeleft = Val(v)
        Case "key_straferight": key_straferight = Val(v)
        Case "key_fire": key_fire = Val(v)
        Case "key_use": key_use = Val(v)
        Case "key_strafe": key_strafe = Val(v)
        Case "key_speed": key_speed = Val(v)
        Case "use_mouse": usemouse = Val(v)
        Case "mouseb_fire": mousebfire = Val(v)
        Case "mouseb_strafe": mousebstrafe = Val(v)
        Case "mouseb_forward": mousebforward = Val(v)
        Case "screenblocks": screenblocks = Val(v)
        Case "detaillevel": detailLevel = Val(v)
        Case "snd_channels": numChannels = Val(v)
        Case "usegamma": usegamma = Val(v)
        Case "chatmacro0" To "chatmacro9"
            If Len(name) = 10 Then chat_macros(Val(Right$(name, 1))) = v
    End Select
End Sub

Private Function DefaultsText() As String
    Dim s As String, i As Long
    s = s & "mouse_sensitivity" & vbTab & mouseSensitivity & vbCrLf
    s = s & "sfx_volume" & vbTab & vbTab & snd_SfxVolume & vbCrLf
    s = s & "music_volume" & vbTab & vbTab & snd_MusicVolume & vbCrLf
    s = s & "show_messages" & vbTab & vbTab & showMessages & vbCrLf
    s = s & "key_right" & vbTab & vbTab & key_right & vbCrLf
    s = s & "key_left" & vbTab & vbTab & key_left & vbCrLf
    s = s & "key_up" & vbTab & vbTab & key_up & vbCrLf
    s = s & "key_down" & vbTab & vbTab & key_down & vbCrLf
    s = s & "key_strafeleft" & vbTab & vbTab & key_strafeleft & vbCrLf
    s = s & "key_straferight" & vbTab & vbTab & key_straferight & vbCrLf
    s = s & "key_fire" & vbTab & vbTab & key_fire & vbCrLf
    s = s & "key_use" & vbTab & vbTab & key_use & vbCrLf
    s = s & "key_strafe" & vbTab & vbTab & key_strafe & vbCrLf
    s = s & "key_speed" & vbTab & vbTab & key_speed & vbCrLf
    s = s & "use_mouse" & vbTab & vbTab & usemouse & vbCrLf
    s = s & "mouseb_fire" & vbTab & vbTab & mousebfire & vbCrLf
    s = s & "mouseb_strafe" & vbTab & vbTab & mousebstrafe & vbCrLf
    s = s & "mouseb_forward" & vbTab & vbTab & mousebforward & vbCrLf
    s = s & "screenblocks" & vbTab & vbTab & screenblocks & vbCrLf
    s = s & "detaillevel" & vbTab & vbTab & detailLevel & vbCrLf
    s = s & "snd_channels" & vbTab & vbTab & numChannels & vbCrLf
    s = s & "usegamma" & vbTab & vbTab & usegamma & vbCrLf
    For i = 0 To 9
        s = s & "chatmacro" & i & vbTab & vbTab & """" & chat_macros(i) & """" & vbCrLf
    Next
    DefaultsText = s
End Function

Public Sub M_SaveDefaults()
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open basedefault For Output As #f
    Print #f, DefaultsText();
    Close #f
End Sub

Public Sub M_LoadDefaults()
    Dim f As Integer, ln As String, nm As String, v As String, p As Long

    ' set everything to base values
    mouseSensitivity = 5
    snd_SfxVolume = 8
    snd_MusicVolume = 8
    showMessages = 1
    key_right = KEY_RIGHTARROW
    key_left = KEY_LEFTARROW
    key_up = KEY_UPARROW
    key_down = KEY_DOWNARROW
    key_strafeleft = Asc(",")
    key_straferight = Asc(".")
    key_fire = KEY_RCTRL
    key_use = Asc(" ")
    key_strafe = KEY_RALT
    key_speed = KEY_RSHIFT
    usemouse = 1
    mousebfire = 0
    mousebstrafe = 1
    mousebforward = 2
    screenblocks = 9
    detailLevel = 0
    numChannels = 8
    usegamma = 0
    chat_macros(0) = HUSTR_CHATMACRO0
    chat_macros(1) = HUSTR_CHATMACRO1
    chat_macros(2) = HUSTR_CHATMACRO2
    chat_macros(3) = HUSTR_CHATMACRO3
    chat_macros(4) = HUSTR_CHATMACRO4
    chat_macros(5) = HUSTR_CHATMACRO5
    chat_macros(6) = HUSTR_CHATMACRO6
    chat_macros(7) = HUSTR_CHATMACRO7
    chat_macros(8) = HUSTR_CHATMACRO8
    chat_macros(9) = HUSTR_CHATMACRO9

    ' read the file in, overriding any set defaults
    If Dir$(basedefault) = "" Then Exit Sub
    On Error GoTo done
    f = FreeFile
    Open basedefault For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        ln = Trim$(Replace(ln, vbTab, " "))
        p = InStr(ln, " ")
        If p > 0 Then
            nm = Left$(ln, p - 1)
            v = Trim$(Mid$(ln, p + 1))
            If Left$(v, 1) = """" Then
                v = Mid$(v, 2)
                If Right$(v, 1) = """" Then v = Left$(v, Len(v) - 1)
            ElseIf LCase$(Left$(v, 2)) = "0x" Then
                v = CStr(CLng("&H" & Mid$(v, 3)))
            End If
            SetDefault nm, v
        End If
    Loop
done:
    Close #f
    If usegamma < 0 Or usegamma > 4 Then usegamma = 0
    If screenblocks < 3 Or screenblocks > 11 Then screenblocks = 9
End Sub

' ------------------------------------------------------------------ screenshots
' Writes screens(0) with the current palette as an 8-bit BMP.
Public Sub M_WriteBMP(ByVal filename As String)
    Dim fh As BMPFILEHEADER, ih As BMPINFOHEADER, pal(0 To 255) As Long
    Dim rows() As Byte, y As Long, i As Long, f As Integer

    For i = 0 To 255
        pal(i) = currentPalette(i * 3 + 2) Or (CLng(currentPalette(i * 3 + 1)) * 256&) Or (CLng(currentPalette(i * 3)) * 65536)
    Next
    ih.biSize = 40
    ih.biWidth = SCREENWIDTH
    ih.biHeight = SCREENHEIGHT                    ' bottom-up
    ih.biPlanes = 1
    ih.biBitCount = 8
    ih.biSizeImage = SCREENSIZE
    ih.biClrUsed = 256
    fh.bfType = &H4D42
    fh.bfOffBits = 14 + 40 + 1024
    fh.bfSize = fh.bfOffBits + SCREENSIZE

    ReDim rows(0 To SCREENSIZE - 1)
    For y = 0 To SCREENHEIGHT - 1
        CopyMemory rows((SCREENHEIGHT - 1 - y) * SCREENWIDTH), screens(y * SCREENWIDTH), SCREENWIDTH
    Next

    On Error Resume Next
    If Dir$(filename) <> "" Then Kill filename
    f = FreeFile
    Open filename For Binary Access Write As #f
    ' BITMAPFILEHEADER is 14 bytes; VB pads the type, so write it field by field
    Put #f, , fh.bfType
    Put #f, , fh.bfSize
    Put #f, , fh.bfReserved1
    Put #f, , fh.bfReserved2
    Put #f, , fh.bfOffBits
    Put #f, , ih
    Put #f, , pal
    Put #f, , rows
    Close #f
End Sub

' M_ScreenShot - finds the next free DOOMnn.BMP
Public Sub M_ScreenShot()
    Dim i As Long, nm As String
    For i = 0 To 99
        nm = DoomExePath & "DOOM" & Format$(i, "00") & ".BMP"
        If Dir$(nm) = "" Then Exit For
    Next
    If i = 100 Then I_Error "M_ScreenShot: Couldn't create a BMP"
    M_WriteBMP nm
    players(consoleplayer).message = "screen shot"
End Sub
