Option  Explicit
Function fCheatSheet(Optional ByVal iRows As Integer = 20, _
                    Optional ByVal iColumns As Integer = 20, _
                    Optional ByVal iVerticalNumber As Integer = 1, _
                    Optional ByVal iHorizontalNumber As Integer = 1, _
                    Optional ByVal bMissingTabs As Boolean = False) _
                    As String

    Dim i As Integer 'counter of verticals
    Dim z As Integer 'counter of horizontals

    For i = iVerticalNumber To (iVerticalNumber + iRows - 1) Step 1
        For z = iHorizontalNumber To (iHorizontalNumber + iColumns - 1) Step 1
            fCheatSheet = fCheatSheet & i * z
            If (z < iHorizontalNumber + iColumns - 1) Then
                If bMissingTabs Then
                fCheatSheet = fCheatSheet & Chr(32)
                Else
                fCheatSheet = fCheatSheet & Chr(9) 'or VbCrLf
                End If
            End If
        Next z
        fCheatSheet = fCheatSheet & vbCrLf
    Next i
    fCheatSheet = fCheatSheet & "vitoshacademy.com"
    
End Function


