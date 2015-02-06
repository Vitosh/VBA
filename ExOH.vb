Option Explicit
Option Compare Text

Public Function ExOH(sInput As String) As Boolean
    
    Dim strArray()  As String
    
    Dim sZero       As String
    Dim sHiks       As String
    
    Dim lHiks       As Long
    Dim lZero       As Long
    Dim lCount      As Long
    
    lCount = Len(sInput)
    ReDim strArray(lCount - 1)
    
    sZero = "o"
    sHiks = "x"
    
    For lCount = 0 To lCount - 1
        strArray(lCount) = Mid(sInput, lCount + 1, 1)
        
        If CStr(strArray(lCount)) = sZero Then
            lZero = lZero + 1
        ElseIf CStr(strArray(lCount)) = sHiks Then
            lHiks = lHiks + 1
        End If
        
    Next lCount
    
    ExOH = (lHiks > 0) And (lHiks = lZero)

End Function
