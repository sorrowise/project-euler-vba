Public Sub PE001()
    Dim res As Long
    Dim i As Integer
    res = 0
    For i = 1 To 1000
        If i Mod 3 = 0 Or i Mod 5 = 0 Then
            res = res + i
        End If
    Next
    Debug.Print res
End Sub
