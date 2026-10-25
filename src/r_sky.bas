Attribute VB_Name = "r_sky"
' r_sky.bas - sky rendering.  The DOOM sky is a texture map like any wall,
' wrapping around.  A 1024 columns equal 360 degrees.  The default sky map is
' 256 columns and repeats 4 times on a 320 screen? (r_sky.c)
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' R_InitSkyMap - called whenever the view size changes.
Public Sub R_InitSkyMap()
    skytexturemid = 100 * FRACUNIT
End Sub
