Option Explicit
Public Function fValidEmail(sTestMail As String) As Boolean

'function code taken from http://www.vbaexpress.com/kb/getarticle.php?kb_id=281
'Thanks to google I do not have to write something like this.

    Dim strArray As Variant
    Dim strItem As Variant
    Dim i As Long, c As String, blnIsItValid As Boolean
    blnIsItValid = True

    i = Len(sTestMail) - Len(Application.Substitute(sTestMail, "@", ""))
    If i <> 1 Then fValidEmail = False: Exit Function
    ReDim strArray(1 To 2)
    strArray(1) = Left(sTestMail, InStr(1, sTestMail, "@", 1) - 1)
    strArray(2) = Application.Substitute(Right(sTestMail, Len(sTestMail) - Len(strArray(1))), "@", "")
    For Each strItem In strArray
        If Len(strItem) <= 0 Then
            blnIsItValid = False
            fValidEmail = blnIsItValid
            Exit Function
        End If
        For i = 1 To Len(strItem)
            c = LCase(Mid(strItem, i, 1))
            If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
                blnIsItValid = False
                fValidEmail = blnIsItValid
                Exit Function
            End If
        Next i
        If Left(strItem, 1) = "." Or Right(strItem, 1) = "." Then
            blnIsItValid = False
            fValidEmail = blnIsItValid
            Exit Function
        End If
    Next strItem
    If InStr(strArray(2), ".") <= 0 Then
        blnIsItValid = False
        fValidEmail = blnIsItValid
        Exit Function
    End If
    i = Len(strArray(2)) - InStrRev(strArray(2), ".")
    If i <> 2 And i <> 3 Then
        blnIsItValid = False
        fValidEmail = blnIsItValid
        Exit Function
    End If
    If InStr(sTestMail, "..") > 0 Then
        blnIsItValid = False
        fValidEmail = blnIsItValid
        Exit Function
    End If
    fValidEmail = blnIsItValid

End Function

Public Sub Main()

    Dim colPeople       As New Collection

    Dim cNewPerson1     As New clsPerson
    Dim cNewPerson2     As New clsPerson
    Dim cNewPerson3     As New clsPerson
    
    Dim Person          As clsPerson
    Dim i               As Integer
    
    colPeople.Add cNewPerson1
    colPeople.Add cNewPerson2
    colPeople.Add cNewPerson3
    
    For Each Person In colPeople
        i = i + 1000
        Person.Name = "Peter" & CStr(i) & sSomeRandomName
        Person.Age = Int(90 * Rnd)
        Person.Email = "review" & CStr(i) & "@vitoshacademy.com"
    Next Person
    
    cNewPerson1.Name = "Vitosh"
    cNewPerson1.Age = 29
    cNewPerson1.Email = "N/A"

    For Each Person In colPeople
        Debug.Print Person.fShowInformation(Person.Name, Person.Age, Person.Email)
    Next Person
    
    Set cNewPerson1 = Nothing
    Set cNewPerson2 = Nothing
    Set cNewPerson3 = Nothing

End Sub


Function sSomeRandomName() As String
'http://stackoverflow.com/questions/22630264/ms-access-visual-basic-generate-random-string-in-text-field
    
    Dim s As String * 8 'fixed length string with 8 characters
    Dim n As Integer
    Dim ch As Integer 'the character
    For n = 1 To Len(s) 'don't hardcode the length twice
        Do
            ch = Rnd() * 127 'This could be more efficient.
            '48 is '0', 57 is '9', 65 is 'A', 90 is 'Z', 97 is 'a', 122 is 'z'.
        Loop While ch < 48 Or ch > 57 And ch < 65 Or ch > 90 And ch < 97 Or ch > 122
        Mid(s, n, 1) = Chr(ch) 'bit more efficient than concatenation
    Next

    sSomeRandomName = s

End Function

