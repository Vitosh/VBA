Option Explicit
Option Compare Text

Public Function ABCheck(sInput As String) As Boolean

    Dim strArray() As String

    Dim sA As String
    Dim sB As String

    Dim lCount As Long

    lCount = Len(sInput)
    ReDim strArray(lCount - 1)
    
    If lCount < 5 Then Exit Function
    
    sA = "a"
    sB = "b"

    For lCount = 0 To lCount - 1
        strArray(lCount) = Mid(sInput, lCount + 1, 1)

        If CStr(strArray(lCount)) = sA Then
            If lCount + 3 <= (Len(sInput) - 1) Then
            strArray(lCount + 4) = Mid(sInput, lCount + 5, 1)
                Debug.Print strArray(lCount + 4)
                Debug.Print Mid(sInput, lCount + 5, 1)
                
                If CStr(strArray(lCount + 4)) = sB Then
                    ABCheck = True
                    Exit Function
                    
                End If
            End If
            
        ElseIf CStr(strArray(lCount)) = sB Then
        
            If lCount + 3 <= (Len(sInput) - 1) Then
            strArray(lCount + 4) = Mid(sInput, lCount + 5, 1)
                Debug.Print strArray(lCount + 4)
                Debug.Print Mid(sInput, lCount + 5, 1)

                If CStr(strArray(lCount + 4)) = sA Then
                    ABCheck = True
                    Exit Function
                End If
            End If
        End If
        
    Next lCount
End Function

