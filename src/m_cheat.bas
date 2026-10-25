Attribute VB_Name = "m_cheat"
' m_cheat.bas - cheat sequence checking (m_cheat.c).
'
' A cheat sequence is the scrambled byte string from the C source (see SCRAMBLE)
' terminated by &HFF.  A sequence containing 1 marks where parameter characters
' are captured (e.g. IDCLEV's two digits).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public Type cheatseq_t
    sequence() As Byte                           ' scrambled bytes, ends with &HFF
    p As Long                                    ' current position, -1 = not started (C: p == 0)
End Type

Private firsttime As Boolean
Private cheat_xlate_table(0 To 255) As Byte
Private initdone As Boolean

Private Function SCRAMBLE(ByVal a As Long) As Long
    SCRAMBLE = ((a And 1) * 128) + ((a And 2) * 32) + (a And 4) + ((a And 8) * 2) _
             + ((a And 16) \ 2) + (a And 32) + ((a And 64) \ 32) + ((a And 128) \ 128)
End Function

' Builds a cheat sequence from plain text, scrambling it exactly like the
' static tables in st_stuff.c.  Chr$(1) marks the start of the parameter and
' each Chr$(0) after it captures one parameter character, so IDCLEV is
' "idclev" & Chr$(1) & Chr$(0) & Chr$(0).
Public Sub cht_MakeSeq(cht As cheatseq_t, ByVal text As String)
    Dim i As Long, c As Long
    ReDim cht.sequence(0 To Len(text))
    For i = 1 To Len(text)
        c = Asc(Mid$(text, i, 1))
        If c = 1 Or c = 0 Then
            cht.sequence(i - 1) = c
        Else
            cht.sequence(i - 1) = SCRAMBLE(c)
        End If
    Next
    cht.sequence(Len(text)) = &HFF
    cht.p = -1
End Sub

' Called in st_stuff module, which handles the input.
' Returns True if the cheat was successful, False if failed.
Public Function cht_CheckCheat(cht As cheatseq_t, ByVal key As Long) As Boolean
    Dim i As Long
    If Not initdone Then
        For i = 0 To 255
            cheat_xlate_table(i) = SCRAMBLE(i)
        Next
        initdone = True
    End If

    If cht.p < 0 Then cht.p = 0                  ' initialize if first time

    key = key And &HFF&
    If cht.sequence(cht.p) = 0 Then
        cht.sequence(cht.p) = key
        cht.p = cht.p + 1
    ElseIf cheat_xlate_table(key) = cht.sequence(cht.p) Then
        cht.p = cht.p + 1
    Else
        cht.p = 0
    End If

    If cht.sequence(cht.p) = 1 Then
        cht.p = cht.p + 1
    ElseIf cht.sequence(cht.p) = &HFF Then       ' end of sequence character
        cht.p = 0
        cht_CheckCheat = True
        Exit Function
    End If
    cht_CheckCheat = False
End Function

' Returns the captured parameter characters (C: cht_GetParam fills a buffer).
Public Function cht_GetParam(cht As cheatseq_t) As String
    Dim i As Long, c As Long, s As String
    i = 0
    Do While cht.sequence(i) <> 1
        i = i + 1
    Loop
    i = i + 1
    Do
        c = cht.sequence(i)
        s = s & Chr$(c)
        cht.sequence(i) = 0
        i = i + 1
    Loop While c <> 0 And cht.sequence(i) <> &HFF
    If cht.sequence(i) = &HFF Then
        ' C appends a 0 terminator here
    End If
    ' strip the trailing terminator if captured
    If Len(s) > 0 Then
        If Right$(s, 1) = Chr$(0) Then s = Left$(s, Len(s) - 1)
    End If
    cht_GetParam = s
End Function
