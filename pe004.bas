Public Function isPalindrome(ByVal n As Long) As Boolean
    Dim numberString, rev As String
    numberString = VBA.Trim(Str(n))
    rev = VBA.StrReverse(numberString)
    isPalindrome = (numberString = rev)
End Function

Public Sub PE004()
    Dim col As New Collection
    Dim i, j, mul, ans As Long
    For i = 999 To 99 Step -1
        For j = 999 To 99 Step -1
            mul = i * j
            If mul Mod 11 = 0 And isPalindrome(mul) Then
                col.Add mul
            End If
        Next j
    Next i
    
    ans = 1
    Dim ele As Variant
    For Each ele In col
        ans = WorksheetFunction.Max(ele, ans)
    Next
    Debug.Print ans
End Sub
