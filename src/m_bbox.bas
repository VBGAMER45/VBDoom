Attribute VB_Name = "m_bbox"
' m_bbox.bas - bounding box helpers (m_bbox.c).
' Boxes are Long arrays indexed BOXTOP, BOXBOTTOM, BOXLEFT, BOXRIGHT.
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public Sub M_ClearBox(box() As Long)
    box(BOXTOP) = MININT
    box(BOXRIGHT) = MININT
    box(BOXBOTTOM) = MAXINT
    box(BOXLEFT) = MAXINT
End Sub

Public Sub M_AddToBox(box() As Long, ByVal x As Long, ByVal y As Long)
    If x < box(BOXLEFT) Then
        box(BOXLEFT) = x
    ElseIf x > box(BOXRIGHT) Then
        box(BOXRIGHT) = x
    End If
    If y < box(BOXBOTTOM) Then
        box(BOXBOTTOM) = y
    ElseIf y > box(BOXTOP) Then
        box(BOXTOP) = y
    End If
End Sub
