Attribute VB_Name = "m_random_c"
' m_random.bas - the deterministic random number table (m_random.c).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public rndtable(0 To 255) As Long                ' filled by InitTables (m_fixed)
Public rndindex As Long
Public prndindex As Long

' Which one is deterministic?
Public Function P_Random() As Long
    prndindex = (prndindex + 1) And &HFF&
    P_Random = rndtable(prndindex)
End Function

Public Function M_Random() As Long
    rndindex = (rndindex + 1) And &HFF&
    M_Random = rndtable(rndindex)
End Function

Public Sub M_ClearRandom()
    rndindex = 0
    prndindex = 0
End Sub
