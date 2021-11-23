Public Sub PE005()
    Dim i, res As Long
    res = 1
    For i = 1 To 20
        res = WorksheetFunction.Lcm(res, i)
    Next
    Debug.Print res
End Sub
