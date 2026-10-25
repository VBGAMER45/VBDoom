VERSION 5.00
Begin VB.Form frmDoom 
   AutoRedraw      =   0   'False
   BackColor       =   &H00000000&
   BorderStyle     =   2  'Sizable
   Caption         =   "DoomVB"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9600
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmDoom"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmDoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmDoom - the game window.  Input is handled by DoomWndProc (i_video.bas).
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Sub Form_Paint()
    I_RepaintWindow
End Sub

Private Sub Form_Resize()
    I_RepaintWindow
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        quitRequested = True
        Cancel = 1
    End If
End Sub
