Public Sub PE009()
    Dim p, a, b, c, n, d As Long
    p = 1000
    For a = 1 To 333
        n = p * p - 2 * p * a
        d = 2 * (p - a)
        If n Mod d = 0 Then
            b = n / d
            c = p - a - b
            Debug.Print (a * b * c)
            Exit Sub
        End If
    Next
End Sub
