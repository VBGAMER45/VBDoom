Attribute VB_Name = "m_argv"
' m_argv.bas - command line parameters (m_argv.c).
' myargv(0) is the program name, as in C.
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public myargc As Long
Public myargv() As String

' Splits Command$ into myargv(), honouring "quoted arguments".
Public Sub M_InitArgs()
    Dim s As String, i As Long, c As String, cur As String
    Dim inq As Boolean, have As Boolean
    ReDim myargv(0 To 0)
    myargv(0) = App.EXEName
    myargc = 1
    s = Command$
    For i = 1 To Len(s) + 1
        If i > Len(s) Then c = " " Else c = Mid$(s, i, 1)
        If c = """" Then
            inq = Not inq
            have = True
        ElseIf (c = " " Or c = vbTab) And Not inq Then
            If have Then
                ReDim Preserve myargv(0 To myargc)
                myargv(myargc) = cur
                myargc = myargc + 1
                cur = ""
                have = False
            End If
        Else
            cur = cur & c
            have = True
        End If
    Next
End Sub

' Returns the argument number (1 to argc-1) or 0 if not present.
Public Function M_CheckParm(ByVal check As String) As Long
    Dim i As Long
    For i = 1 To myargc - 1
        If StrComp(check, myargv(i), vbTextCompare) = 0 Then
            M_CheckParm = i
            Exit Function
        End If
    Next
    M_CheckParm = 0
End Function
