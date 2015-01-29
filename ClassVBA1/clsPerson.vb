Option Explicit

Public sName    As String
Public iAge     As Integer
Public sEmail   As String

Public Property Get Name() As String

    Name = sName
    Debug.Print "Name taken from memory."
    
End Property

Public Property Let Name(Value As String)
    
    If Len(CStr(Value)) Then
        sName = Value
        Debug.Print "Name is assigned successfully - " & Value
    Else
        Debug.Print "The name cannot be empty!"
        sName = ""
    End If
    
End Property

Public Property Get Age() As Integer

    Age = iAge
    Debug.Print "Age taken from memory."
    
End Property

Public Property Let Age(Value As Integer)

    If (1 < Value) And (Value < 100) Then
        iAge = Value
        Debug.Print "Age is assigned successfully - " & Value
    Else
        Debug.Print "The value should be between 1 and 100!"
    End If
    
End Property

Public Property Get Email() As String

    Email = sEmail
    Debug.Print "Email taken from memory."
    
End Property

Public Property Let Email(Value As String)
    
    If fValidEmail(Value) Then
        sEmail = Value
        Debug.Print "Email is assigned successfully - " & Value
    ElseIf StrComp(Value, "N/A") = 0 Then
        sEmail = ""
    Else
        Debug.Print "Please, enter a valid e-mail!"
    End If
    
End Property

Public Function fShowInformation(fShowName As String, fShowAge As Integer, fShowMail As String) As String
           
    If Len(fShowMail) Then
        fShowInformation = fShowName & " is " & fShowAge & " years old and his e-mail address is " & fShowMail
    Else
        fShowInformation = fShowName & " is " & fShowAge & " years old and no email is available!"
    End If
    
End Function

Private Sub Class_Initialize()
    Debug.Print "ClsPerson is initialized!"
End Sub

Private Sub Class_Terminate()
    Debug.Print "ClsPerson is terminated for " & sName
End Sub
