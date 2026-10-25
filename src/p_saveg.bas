Attribute VB_Name = "p_saveg"
' p_saveg.bas - archiving: SaveGame I/O (p_saveg.c).
'
' The C version serialised each structure by hand and cleared pointers such
' as mobj->target on load.  Here every reference is an array index, so the
' arrays are written as they are (VB Put/Get on a file opened For Binary by
' g_game) and all references survive a save/load.  The format is specific to
' this port.
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' P_ArchivePlayers
Public Sub P_ArchivePlayers(ByVal f As Integer)
    Put #f, , players
End Sub

' P_UnArchivePlayers
Public Sub P_UnArchivePlayers(ByVal f As Integer)
    Get #f, , players
End Sub

' P_ArchiveWorld
Public Sub P_ArchiveWorld(ByVal f As Integer)
    Put #f, , numsectors
    Put #f, , numlines
    Put #f, , numsides
    Put #f, , sectors
    Put #f, , lines
    Put #f, , sides
End Sub

' P_UnArchiveWorld
Public Sub P_UnArchiveWorld(ByVal f As Integer)
    Dim ns As Long, nl As Long, nsd As Long
    Get #f, , ns
    Get #f, , nl
    Get #f, , nsd
    If ns <> numsectors Or nl <> numlines Or nsd <> numsides Then
        I_Error "P_UnArchiveWorld: savegame does not match the level"
    End If
    Get #f, , sectors
    Get #f, , lines
    Get #f, , sides
End Sub

' Thinkers
Public Sub P_ArchiveThinkers(ByVal f As Integer)
    Dim n As Long
    n = maxthinkers
    Put #f, , n
    Put #f, , thinkers
    Put #f, , mobjs
    Put #f, , ceilings
    Put #f, , doors
    Put #f, , floormoves
    Put #f, , plats
    Put #f, , fireflickers
    Put #f, , lightflashes
    Put #f, , strobes
    Put #f, , glows
    n = UBound(blocklinks)
    Put #f, , n
    Put #f, , blocklinks
End Sub

' P_UnArchiveThinkers
Public Sub P_UnArchiveThinkers(ByVal f As Integer)
    Dim n As Long

    Get #f, , n
    P_SetThinkerCapacity n
    Get #f, , thinkers
    Get #f, , mobjs
    Get #f, , ceilings
    Get #f, , doors
    Get #f, , floormoves
    Get #f, , plats
    Get #f, , fireflickers
    Get #f, , lightflashes
    Get #f, , strobes
    Get #f, , glows
    P_RebuildFreeList

    Get #f, , n
    If n <> UBound(blocklinks) Then I_Error "P_UnArchiveThinkers: savegame does not match the level"
    Get #f, , blocklinks
End Sub

' P_ArchiveSpecials
Public Sub P_ArchiveSpecials(ByVal f As Integer)
    Put #f, , activeceilings
    Put #f, , activeplats
    Put #f, , buttonlist
    Put #f, , levelTimer
    Put #f, , levelTimeCount
    Put #f, , iquehead
    Put #f, , iquetail
    Put #f, , itemrespawnque
    Put #f, , itemrespawntime
    Put #f, , totalkills
    Put #f, , totalitems
    Put #f, , totalsecret
End Sub

' P_UnArchiveSpecials
Public Sub P_UnArchiveSpecials(ByVal f As Integer)
    Get #f, , activeceilings
    Get #f, , activeplats
    Get #f, , buttonlist
    Get #f, , levelTimer
    Get #f, , levelTimeCount
    Get #f, , iquehead
    Get #f, , iquetail
    Get #f, , itemrespawnque
    Get #f, , itemrespawntime
    Get #f, , totalkills
    Get #f, , totalitems
    Get #f, , totalsecret
End Sub
