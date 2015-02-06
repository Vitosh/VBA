Option Explicit
Public Const LETTERS_IN_ALPHABET = 26

Public Function caesar_encrypt(sInput As String, n As Integer) As String

    Dim sResult                 As String
    Dim i                       As Integer
    Dim iAsciiInput             As Integer
    Dim iAsciiOutput             As Integer
    
    sResult = sInput
    
    For i = 0 To Len(sResult) - 1
            
        iAsciiInput = Asc(Mid(sInput, i + 1, 1))
        iAsciiOutput = Asc(Mid(sInput, i + 1, 1))
        
        'Capital letters
        If iAsciiInput > 64 And iAsciiInput < 91 Then
        Mid(sResult, i + 1, 1) = Chr(((iAsciiOutput + n - 65) Mod LETTERS_IN_ALPHABET) + 65)
        
        'Small letters
        ElseIf iAsciiInput > 96 And iAsciiInput < 123 Then
        Mid(sResult, i + 1, 1) = Chr(((iAsciiOutput + n - 97) Mod LETTERS_IN_ALPHABET) + 97)
        
        Else
        'Others
        Mid(sResult, i + 1, 1) = Mid(sInput, i + 1, 1)
        
        End If
    Next i

    caesar_encrypt = sResult
    
End Function
