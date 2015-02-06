Option Explicit

Public Function fill_tetrahedron(ByVal num As Integer) As Double

    fill_tetrahedron = Round((num ^ 3) / (6 * Sqr(2)) / 1000, 2)
    
End Function

Public Function tetrahedron_filled(iWater As Integer, ParamArray arrNumbers() As Variant) As Integer
    'Pay attention for the way the parameters are put.
    'First we put the iWater and then the tetrahedrons.
    
    Dim arrResult()         As Variant
    Dim iResult             As Long
    
    Dim z                   As Long
    Dim i                   As Long
    
    ReDim arrResult(UBound(arrNumbers))
    
    For i = LBound(arrNumbers) To UBound(arrNumbers) Step 1
        arrResult(i) = CInt(arrNumbers(i))
    Next i
    
    InsertionSort arrResult(), LBound(arrResult), UBound(arrResult)
    
    i = 0
    
    For i = LBound(arrResult) To UBound(arrResult) Step 1
        z = CLng(arrResult(i))
        iWater = iWater - fill_tetrahedron(z)

        If iWater < 0 Then
            tetrahedron_filled = i
            Exit Function
        End If
    Next i
    
    tetrahedron_filled = UBound(arrResult) + 1
        
End Function

Public Sub InsertionSort(ByRef a(), ByVal lo0 As Long, ByVal hi0 As Long)
    Dim i As Long, j As Long, v As Long

    For i = lo0 + 1 To hi0
        v = a(i)
        j = i
        Do While j > lo0
            If Not a(j - 1) > v Then Exit Do
            a(j) = a(j - 1)
            j = j - 1
        Loop
        a(j) = v
    Next i
End Sub

