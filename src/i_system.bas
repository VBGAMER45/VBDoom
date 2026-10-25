Attribute VB_Name = "i_system"
' i_system.bas - system specific interface: timing, errors, quitting, logging
' (i_system.c, i_main.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private basetime As Long
Private timeinit As Boolean
Private inerror As Boolean
Private logopen As Boolean
Private logfile As Integer

Public quitRequested As Boolean                  ' set by the window procedure on WM_CLOSE
Public dbgPhase As String                        ' breadcrumb for crash reports

' Program entry point (i_main.c)
Public Sub Main()
    DoomExePath = App.Path
    If Right$(DoomExePath, 1) <> "\" Then DoomExePath = DoomExePath & "\"
    M_InitArgs
    timeBeginPeriod 1
    D_DoomMain
End Sub

' returns time in 1/35th second tics
Public Function I_GetTime() As Long
    Dim t As Long
    t = timeGetTime()
    If Not timeinit Then
        basetime = t
        timeinit = True
    End If
    I_GetTime = Int(CDbl(t - basetime) * TICRATE / 1000#)
End Function

' milliseconds since startup (for the sound/music code)
Public Function I_GetTimeMS() As Long
    If Not timeinit Then I_GetTime
    I_GetTimeMS = timeGetTime() - basetime
End Function

Public Sub I_Init()
    I_InitSound
    I_InitMusic
End Sub

' wait for vertical retrace or pause a bit (count is in 1/70 s)
Public Sub I_WaitVBL(ByVal count As Long)
    Sleep (count * 1000) \ 70
End Sub

Public Sub I_BeginRead()
End Sub

Public Sub I_EndRead()
End Sub

Public Sub I_Printf(ByVal s As String)
    On Error Resume Next
    If Not logopen Then
        logfile = FreeFile
        Open DoomExePath & "doomvb.log" For Output As #logfile
        logopen = True
    End If
    Print #logfile, s
    Debug.Print s
End Sub

Private Sub I_CloseLog()
    On Error Resume Next
    If logopen Then
        Close #logfile
        logopen = False
    End If
End Sub

Private Sub I_ShutdownAll()
    On Error Resume Next
    I_ShutdownSound
    I_ShutdownMusic
    I_ShutdownNetwork
    I_ShutdownGraphics
    timeEndPeriod 1
End Sub

Public Sub I_Quit()
    On Error Resume Next
    D_QuitNetGame
    M_SaveDefaults
    I_ShutdownAll
    I_CloseLog
    End
End Sub

Public Sub I_Error(ByVal msg As String)
    If inerror Then End
    inerror = True
    On Error Resume Next
    I_Printf "Error: " & msg
    If Len(dbgPhase) > 0 Then I_Printf "  (while in " & dbgPhase & ")"

    ' Shutdown. Here might be other errors.
    If demorecording Then G_CheckDemoStatus
    D_QuitNetGame
    I_ShutdownAll
    I_CloseLog
    If M_CheckParm("-hidden") = 0 Then
        MsgBox msg, vbCritical, "DoomVB"
    End If
    End
End Sub

' Converts a trapped VB run time error into an I_Error with context.
Public Sub I_RuntimeError(ByVal where As String)
    I_Error "Run time error " & Err.Number & " (" & Err.Description & ") in " & where & _
        IIf(Erl <> 0, " line " & Erl, "") & " at gametic " & gametic
End Sub
