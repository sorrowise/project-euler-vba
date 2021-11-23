Public Sub PE003()
    Dim i, n As Longlong
    i = 2
    n = 600851475143#
    While i ^ 2 < n
        While n Mod i = 0
            n = n / i
        Wend
        If i > 2 Then
            i = i + 2
        Else
            i = i + 1
        End If
    Wend
    Debug.Print n
End Sub
