Attribute VB_Name = "m_fixed"
' m_fixed.bas - fixed point arithmetic, shift helpers and the trig tables
' (m_fixed.c, tables.c).  The tables are the original values, loaded from
' the embedded resource res/tables.bin.
'
' DoomVB: a Visual Basic 6 port of the DOOM source release.
' Copyright (C) 1993-1996 id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public finesine() As Long                        ' 0..10239 (5 * FINEANGLES / 4)
Public finecosine() As Long                      ' 0..8191  (C: &finesine[FINEANGLES/4])
Public finetangent() As Long                     ' 0..4095 (+ 4096..8191 = finesine, as laid out in C memory)
Public tantoangle() As Long                      ' 0..2048

Public pow2(0 To 31) As Long                     ' pow2(n) = 1 << n  (pow2(31) = &H80000000)

' (a * b) >> 16 with the exact 64-bit intermediate of the C version.
' a = ah * 65536 + al  ->  (a * b) >> 16 = ah * b + ((al * b) >> 16)
' ah * b deliberately wraps (integer overflow checks are off in the build).
Public Function FixedMul(ByVal a As Long, ByVal b As Long) As Long
    Dim ah As Long
    ah = (a And &HFFFF0000) \ &H10000
    FixedMul = ah * b + Int(CDbl(a And &HFFFF&) * CDbl(b) / 65536#)
End Function

' (a << 16) / b with the overflow clamp of the C version.
Public Function FixedDiv(ByVal a As Long, ByVal b As Long) As Long
    If Int(Abs(CDbl(a)) / 16384#) >= Abs(CDbl(b)) Then
        If (a Xor b) < 0 Then FixedDiv = MININT Else FixedDiv = MAXINT
        Exit Function
    End If
    FixedDiv = FixedDiv2(a, b)
End Function

' Exact truncating integer division (a << 16) / b, like the idiv of the
' original DOS assembly version.  A double gives the estimate, the exactly
' computed remainder corrects it, so the result does not depend on how the
' compiler rounds floating point intermediates.
Public Function FixedDiv2(ByVal a As Long, ByVal b As Long) As Long
    Dim na As Double, ba As Double, q As Double, r As Double

    na = Abs(CDbl(a)) * 65536#                   ' exact, < 2^48
    ba = Abs(CDbl(b))
    If ba = 0 Then I_Error "FixedDiv: divide by zero"
    q = Int(na / ba)
    r = na - q * ba                              ' exact
    Do While r < 0
        q = q - 1
        r = r + ba
    Loop
    Do While r >= ba
        q = q + 1
        r = r - ba
    Loop
    If (a Xor b) < 0 Then q = -q
    If q >= 2147483648# Or q < -2147483648# Then I_Error "FixedDiv: divide by zero"
    FixedDiv2 = q
End Function

' x >> n, arithmetic (sign-propagating) shift, 0 <= n <= 31
Public Function SAR(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then
        SAR = x
    ElseIf n = 31 Then
        If x < 0 Then SAR = -1 Else SAR = 0
    Else
        SAR = (x And Not (pow2(n) - 1)) \ pow2(n)
    End If
End Function

' x >> n, logical (unsigned) shift, 0 <= n <= 31
Public Function SHR(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then
        SHR = x
    ElseIf n = 31 Then
        If x < 0 Then SHR = 1 Else SHR = 0
    Else
        SHR = (x And &H7FFFFFFF) \ pow2(n)
        If x < 0 Then SHR = SHR Or pow2(31 - n)
    End If
End Function

' x << n (wraps like C), 0 <= n <= 31
Public Function SHL(ByVal x As Long, ByVal n As Long) As Long
    If n = 31 Then
        If (x And 1) <> 0 Then SHL = &H80000000 Else SHL = 0
    Else
        SHL = x * pow2(n)                        ' wraps (overflow checks are off)
    End If
End Function

' x >> FRACBITS (floor), the most common shift
Public Function FIXTOINT(ByVal x As Long) As Long
    FIXTOINT = (x And &HFFFF0000) \ &H10000
End Function

' wrap-safe angle arithmetic (also safe when overflow checks are enabled)
Public Function AngAdd(ByVal a As Long, ByVal b As Long) As Long
    Dim d As Double
    d = CDbl(a) + CDbl(b)
    If d > 2147483647# Then
        d = d - 4294967296#
    ElseIf d < -2147483648# Then
        d = d + 4294967296#
    End If
    AngAdd = d
End Function

Public Function AngSub(ByVal a As Long, ByVal b As Long) As Long
    Dim d As Double
    d = CDbl(a) - CDbl(b)
    If d > 2147483647# Then
        d = d - 4294967296#
    ElseIf d < -2147483648# Then
        d = d + 4294967296#
    End If
    AngSub = d
End Function

' unsigned 32-bit value of a Long, as a Double
Public Function UDbl(ByVal a As Long) As Double
    If a < 0 Then UDbl = CDbl(a) + 4294967296# Else UDbl = a
End Function

' Double (any range) -> Long, wrapping modulo 2^32 like a C cast of an integer
Public Function WrapLong(ByVal d As Double) As Long
    d = d - 4294967296# * Int((d + 2147483648#) / 4294967296#)
    WrapLong = d
End Function

Public Sub InitTables()
    Dim b() As Byte, i As Long, ofs As Long

    For i = 0 To 30
        pow2(i) = 2 ^ i
    Next
    pow2(31) = &H80000000

    b = LoadResData(101, "CUSTOM")
    ReDim finetangent(0 To 8191)
    ReDim finesine(0 To 10239)
    ReDim finecosine(0 To 8191)
    ReDim tantoangle(0 To 2048)
    ofs = 0
    CopyMemory finetangent(0), b(ofs), 4096& * 4: ofs = ofs + 4096& * 4
    CopyMemory finesine(0), b(ofs), 10240& * 4: ofs = ofs + 10240& * 4
    CopyMemory tantoangle(0), b(ofs), 2049& * 4: ofs = ofs + 2049& * 4
    For i = 0 To 255
        rndtable(i) = b(ofs + i)
    Next
    ofs = ofs + 256
    For i = 0 To 1279
        gammatable(i \ 256, i And 255) = b(ofs + i)
    Next
    For i = 0 To 8191
        finecosine(i) = finesine(i + 2048)
    Next
    For i = 0 To 4095
        finetangent(4096 + i) = finesine(i)
    Next
End Sub

' tables.c: SlopeDiv  (num and den are unsigned in C; num << 3 wraps)
Public Function SlopeDiv(ByVal num As Long, ByVal den As Long) As Long
    Dim ans As Long, n8 As Double
    If den >= 0 And den < 512 Then
        SlopeDiv = SLOPERANGE
        Exit Function
    End If
    If num >= 0 And num < &H20000000 And den > 0 Then
        ans = (num * 8) \ (den \ 256)
    Else
        n8 = UDbl(num) * 8#
        n8 = n8 - 4294967296# * Int(n8 / 4294967296#)
        n8 = Int(n8 / Int(UDbl(den) / 256#))
        If n8 > SLOPERANGE Then n8 = SLOPERANGE
        ans = n8
    End If
    If ans <= SLOPERANGE And ans >= 0 Then SlopeDiv = ans Else SlopeDiv = SLOPERANGE
End Function
