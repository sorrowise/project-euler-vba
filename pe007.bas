Public Function primesieve(n As Long) As Variant
    Dim i, j, arr() As Long
    ReDim arr(n)
    Dim primes As New Collection
    For i = 1 To n: arr(i) = 1: Next
    For i = 2 To Int(n ^ 0.5)
        If arr(i) = 1 Then
            For j = 2 * i To n Step i
                arr(j) = 0
            Next j
        End If
    Next i
    For i = 2 To n
        If arr(i) = 1 Then
            primes.Add i
        End If
    Next i
    primesieve = collectiontoarray(primes)
End Function


Public Function nthprime(n As Long) As Long
    If n <= 3 Then
        Dim sp As Variant
        sp = Array(2, 3, 5)
        nthprime = sp(n)
    Else
        Dim u As Long
        Dim arr As Variant
        u = Int(n * Log(n * Log(n))) + 1
        arr = primesieve(u)
        nthprime = arr(n)
    End If
End Function


Public Sub PE007()
    Debug.Print nthprime(10001)
End Sub
