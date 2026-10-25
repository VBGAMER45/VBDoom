Attribute VB_Name = "w_wad"
' w_wad.bas - WAD file loading and lump access (w_wad.c).
'
' Every WAD file is read completely into wadbuf().  A "cached lump" is simply
' the byte offset of the lump inside wadbuf(): W_CacheLumpNum / W_CacheLumpName
' return that offset and the Rd* helpers read little-endian values from it.
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public Type lumpinfo_t
    name As String                               ' upper case, no trailing NULs
    position As Long                             ' offset into wadbuf()
    size As Long
End Type

Public wadbuf() As Byte
Public lumpinfo() As lumpinfo_t
Public numlumps As Long

Private lumpindex As Collection                  ' name -> lump number (last one wins)
Private wadbufsize As Long

Public Function RdByte(ByVal ofs As Long) As Long
    RdByte = wadbuf(ofs)
End Function

' signed 16-bit
Public Function RdShort(ByVal ofs As Long) As Long
    Dim v As Long
    v = wadbuf(ofs) Or (CLng(wadbuf(ofs + 1)) * 256&)
    If v >= 32768 Then v = v - 65536
    RdShort = v
End Function

' unsigned 16-bit
Public Function RdUShort(ByVal ofs As Long) As Long
    RdUShort = wadbuf(ofs) Or (CLng(wadbuf(ofs + 1)) * 256&)
End Function

' signed 32-bit
Public Function RdLong(ByVal ofs As Long) As Long
    Dim v As Long
    CopyMemory v, wadbuf(ofs), 4
    RdLong = v
End Function

' 8 character lump/texture name, upper case, NUL trimmed
Public Function RdName(ByVal ofs As Long) As String
    Dim i As Long, s As String, c As Long
    For i = 0 To 7
        c = wadbuf(ofs + i)
        If c = 0 Then Exit For
        s = s & Chr$(c)
    Next
    RdName = UCase$(s)
End Function

Private Sub AddFile(ByVal filename As String)
    Dim f As Integer, flen As Long, ident As String
    Dim nlumps As Long, tableofs As Long, base As Long
    Dim i As Long, ofs As Long, startlump As Long

    If Dir$(filename) = "" Then
        I_Printf " couldn't open " & filename
        Exit Sub
    End If
    I_Printf " adding " & filename

    f = FreeFile
    Open filename For Binary Access Read As #f
    flen = LOF(f)
    base = wadbufsize
    wadbufsize = wadbufsize + flen
    If base = 0 Then
        ReDim wadbuf(0 To wadbufsize - 1)
    Else
        ReDim Preserve wadbuf(0 To wadbufsize - 1)
    End If
    If flen > 0 Then
        Dim tmp() As Byte
        ReDim tmp(0 To flen - 1)
        Get #f, 1, tmp
        CopyMemory wadbuf(base), tmp(0), flen
        Erase tmp
    End If
    Close #f

    startlump = numlumps
    If UCase$(Right$(filename, 4)) <> ".WAD" Then
        ' single lump file
        ReDim Preserve lumpinfo(0 To numlumps)
        lumpinfo(numlumps).name = UCase$(Left$(ExtractBaseName(filename), 8))
        lumpinfo(numlumps).position = base
        lumpinfo(numlumps).size = flen
        numlumps = numlumps + 1
    Else
        ident = Chr$(wadbuf(base)) & Chr$(wadbuf(base + 1)) & Chr$(wadbuf(base + 2)) & Chr$(wadbuf(base + 3))
        If ident <> "IWAD" And ident <> "PWAD" Then
            I_Error "Wad file " & filename & " doesn't have IWAD or PWAD id"
        End If
        If ident = "PWAD" Then modifiedgame = True
        nlumps = RdLong(base + 4)
        tableofs = RdLong(base + 8)
        ReDim Preserve lumpinfo(0 To numlumps + nlumps - 1)
        For i = 0 To nlumps - 1
            ofs = base + tableofs + i * 16
            lumpinfo(numlumps + i).position = base + RdLong(ofs)
            lumpinfo(numlumps + i).size = RdLong(ofs + 4)
            lumpinfo(numlumps + i).name = RdName(ofs + 8)
        Next
        numlumps = numlumps + nlumps
    End If

    For i = startlump To numlumps - 1
        On Error Resume Next
        lumpindex.Remove lumpinfo(i).name
        On Error GoTo 0
        lumpindex.Add i, lumpinfo(i).name
    Next
End Sub

Private Function ExtractBaseName(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(path, "\")
    If p > 0 Then path = Mid$(path, p + 1)
    p = InStrRev(path, ".")
    If p > 0 Then path = Left$(path, p - 1)
    ExtractBaseName = path
End Function

' Pass a list of file names.  All files are optional, but at least one file
' must be found.  Lump names can appear multiple times; the last one wins.
Public Sub W_InitMultipleFiles(files() As String, ByVal count As Long)
    Dim i As Long
    numlumps = 0
    wadbufsize = 0
    Set lumpindex = New Collection
    For i = 0 To count - 1
        AddFile files(i)
    Next
    If numlumps = 0 Then I_Error "W_InitFiles: no files found"
End Sub

' Returns -1 if name not found.
Public Function W_CheckNumForName(ByVal name As String) As Long
    Dim v As Variant
    name = UCase$(name)
    If Len(name) > 8 Then name = Left$(name, 8)
    On Error GoTo notfound
    v = lumpindex(name)
    W_CheckNumForName = v
    Exit Function
notfound:
    W_CheckNumForName = -1
End Function

' Calls W_CheckNumForName, but bombs out if not found.
Public Function W_GetNumForName(ByVal name As String) As Long
    Dim i As Long
    i = W_CheckNumForName(name)
    If i = -1 Then I_Error "W_GetNumForName: " & name & " not found!"
    W_GetNumForName = i
End Function

Public Function W_LumpLength(ByVal lump As Long) As Long
    If lump < 0 Or lump >= numlumps Then I_Error "W_LumpLength: " & lump & " >= numlumps"
    W_LumpLength = lumpinfo(lump).size
End Function

' Copies the lump into dest() (resized to the lump length).
Public Sub W_ReadLump(ByVal lump As Long, dest() As Byte)
    Dim n As Long
    If lump < 0 Or lump >= numlumps Then I_Error "W_ReadLump: " & lump & " >= numlumps"
    n = lumpinfo(lump).size
    If n = 0 Then
        Erase dest
        Exit Sub
    End If
    ReDim dest(0 To n - 1)
    CopyMemory dest(0), wadbuf(lumpinfo(lump).position), n
End Sub

' Returns the offset of the lump data inside wadbuf().
Public Function W_CacheLumpNum(ByVal lump As Long) As Long
    If lump < 0 Or lump >= numlumps Then I_Error "W_CacheLumpNum: " & lump & " >= numlumps"
    W_CacheLumpNum = lumpinfo(lump).position
End Function

Public Function W_CacheLumpName(ByVal name As String) As Long
    W_CacheLumpName = W_CacheLumpNum(W_GetNumForName(name))
End Function

Public Function W_LumpName(ByVal lump As Long) As String
    W_LumpName = lumpinfo(lump).name
End Function
