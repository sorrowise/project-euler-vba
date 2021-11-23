'approach 1

Public Sub PE002()
    Dim fib As New Collection
    Dim i, res As Long
    i = 3
    res = 0
    fib.Add 1
    fib.Add 2
    fib.Add 3
    While fib(i) <= 4000000
        If fib(i) Mod 2 = 0 Then
            res = res + fib(i)
        End If
        i = i + 1
        fib.Add fib(i - 1) + fib(i - 2)
    Wend
    Debug.Print res + 2
End Sub
