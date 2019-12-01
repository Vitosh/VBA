Option Explicit

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub Main()
    
    Dim url As String: url = "https://www.mrci.com/ohlc/ohlc-all.php"
    Dim target As String: target = "Brent Crude Oil(ICE)"
    
    Dim appIE As Object
    Set appIE = StartIE(url, appIE)

    Dim allRowsOfData As Variant
    allRowsOfData = appIE.document.getElementsByClassName("strat")
    
    Dim headers As Variant
    headers = GenerateHeaders(allRowsOfData)
    
    Dim title As String
    title = GenerateTitle(appIE)
    
    Dim targetPeriods As Object
    Set targetPeriods = CreateObject("Scripting.Dictionary")
    Set targetPeriods = GenerateTargetPeriods(allRowsOfData, headers, target, targetPeriods)
    
    WriteToExcel targetPeriods, headers, target, title
    
    appIE.Quit

End Sub

Public Sub WriteToExcel(targetPeriods As Object, headers As Variant, target As String, title As String)
    
    Dim row As Long: row = 1
    Dim col As Long: col = 1
    Dim element As Variant
    Dim unit As Long
    Dim child As Variant
    
    With Worksheets(1)
        .Cells.Delete
        .Cells(row, col) = title
        row = row + 1
        .Cells(row, col) = target
        row = row + 2
        
        For Each element In headers
            .Cells(row, col) = element
            col = col + 1
        Next
        row = row + 1
        
        For Each child In targetPeriods
            col = 1
            For unit = 1 To targetPeriods(child).Count
                .Cells(row, col) = targetPeriods(child)(unit)
                col = col + 1
            Next
            row = row + 1
        Next child
        
        Dim startingCol As String
        Dim endingCol As String
        startingCol = NumberToLetters(2)
        endingCol = NumberToLetters(UBound(headers) + 1)
        
        .Columns(startingCol & ":" & endingCol).EntireColumn.AutoFit
    End With
    
End Sub

Public Function GenerateTitle(appIE As Object) As String
    
    Dim titleObj As Variant
    titleObj = appIE.document.getElementsByClassName("title1")
    GenerateTitle = titleObj.innerText
    
End Function

Public Function GenerateTargetPeriods(allRowsOfData As Variant, headers As Variant, target As String, targetPeriods As Object) As Variant

    Dim child As Variant
    Dim child2 As Variant
    Dim myKey As Variant
    Dim i As Long
    Dim found As Boolean: found = False
    
    For Each child In allRowsOfData.Children
        For Each child2 In child.Children
        
            If InStr(1, child2.innerText, target) Then found = True
            If InStr(1, child2.outerhtml, "th class=") And (Not InStr(1, child2.outerhtml, target) > 0) Then
                found = False
            End If
            
            If found And child2.Cells.Length = UBound(headers) + 1 Then
                
                myKey = target & child2.Children(0).innerText
                targetPeriods.Add myKey, New Collection
                
                For i = LBound(headers) To UBound(headers)
                    targetPeriods(myKey).Add (child2.Children(i).innerText)
                    Debug.Print child2.Children(i).innerText
                Next
                
            End If
        Next
    Next
    
    Set GenerateTargetPeriods = targetPeriods

End Function

Public Function GenerateHeaders(allRowsOfData As Variant) As Variant

    Dim headers As Variant
    headers = Split(allRowsOfData.Rows(2).innerText, vbCrLf)
    headers = RemoveEmptyElementsFromArray(headers)
    GenerateHeaders = headers
    
End Function

Public Function StartIE(url As String, appIE As Object) As Object
    
    Set appIE = CreateObject("InternetExplorer.Application")
    With appIE
        .Navigate url
        .Visible = True
    End With

    WaitSomeMilliseconds 2000
    Do While appIE.Busy: DoEvents: Loop
    
    Set StartIE = appIE
    
End Function

Public Sub WaitSomeMilliseconds(Optional Milliseconds As Long = 1000)
    Sleep Milliseconds
End Sub

Public Function RemoveEmptyElementsFromArray(myArray As Variant) As Variant
    
    Dim i As Long, j As Long
    ReDim newArray(LBound(myArray) To UBound(myArray))
    
    For i = LBound(myArray) To UBound(myArray)
        If Trim(myArray(i)) <> "" Then
            j = j + 1
            newArray(j) = myArray(i)
        End If
    Next i
    
    ReDim Preserve newArray(LBound(myArray) To j - 1)
    RemoveEmptyElementsFromArray = newArray
    
End Function

Public Function NumberToLetters(number As Long) As String
    NumberToLetters = Split(Cells(1, number).Address, "$")(1)
End Function
