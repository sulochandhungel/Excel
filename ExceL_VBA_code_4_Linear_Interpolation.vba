' Linear Interpolation Function
'
Function LinInterp(X, InArray, OutArray)
'Application.Volatile
    Num = Application.CountA(OutArray)
    NumIn = Application.CountA(InArray)
    For i = 2 To NumIn
        If InArray(i) < InArray(i - 1) Then
            LinInterp = "Sort Error"
           Exit Function 'Input array must be sorted from smallest to largest.
        End If
    Next i
    For i = 1 To Num
        If X <= InArray(i) Then
            If i = 1 Then
                LinInterp = OutArray(1)
                Exit Function
            End If
            ax = X - InArray(i - 1)
            bx = InArray(i) - InArray(i - 1)
            by = OutArray(i) - OutArray(i - 1)
            LinInterp = OutArray(i - 1) + (ax / bx) * by
            Exit Function
        End If
    Next i
'   LinInterp = ">Table"  (Rating table exceeded - linearly extrapolate)
    ax = X - InArray(Num)
    bx = InArray(Num) - InArray(Num - 1)
    by = OutArray(Num) - OutArray(Num - 1)
    LinInterp = OutArray(Num) + (ax / bx) * by
End Function

