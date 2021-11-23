Public Function stringProduct(ByVal s As String) As LongLong
    Dim i, res As LongLong
    res = 1
    For i = 1 To Len(s)
        res = res * CLng(Mid(s, i, 1))
    Next
    stringProduct = res
End Function

Public Sub PE008()
    Dim s, number As String
    Dim i, ans As LongLong
    
    number = "73167176531330624919225119674426574742355349194934" & _
            "96983520312774506326239578318016984801869478851843" & _
            "85861560789112949495459501737958331952853208805511" & _
            "12540698747158523863050715693290963295227443043557" & _
            "66896648950445244523161731856403098711121722383113" & _
            "62229893423380308135336276614282806444486645238749" & _
            "30358907296290491560440772390713810515859307960866" & _
            "70172427121883998797908792274921901699720888093776" & _
            "65727333001053367881220235421809751254540594752243" & _
            "52584907711670556013604839586446706324415722155397" & _
            "53697817977846174064955149290862569321978468622482" & _
            "83972241375657056057490261407972968652414535100474" & _
            "82166370484403199890008895243450658541227588666881" & _
            "16427171479924442928230863465674813919123162824586" & _
            "17866458359124566529476545682848912883142607690042" & _
            "24219022671055626321111109370544217506941658960408" & _
            "07198403850962455444362981230987879927244284909188" & _
            "84580156166097919133875499200524063689912560717606" & _
            "05886116467109405077541002256983155200055935729725" & _
            "71636269561882670428252483600823257530420752963450"
    ans = 1
    For i = 1 To 997
        s = Mid(number, i, 13)
        ans = WorksheetFunction.Max(ans, stringProduct(s))
    Next
    
    Debug.Print ans
        
End Sub