Option Explicit

Sub Main()
      
    Dim theWord As String: theWord = "OPABRACADABRAK"
    Dim length As Long: length = Len(theWord)
    Dim maxLength As Long: maxLength = 1
    Dim startAt As Long: startAt = 1
    
    ClearTable
    EditTheTable length
    WriteTheWord theWord
    
    ReDim matrix(length - 1, length - 1) As Long
    
    'For 1:
    Dim i As Long
    For i = LBound(matrix) To UBound(matrix)
        tblMatrix.Cells(i + 1, i + 1).Interior.Color = vbYellow
        tblMatrix.Cells(i + 1, i + 1) = 1
    Next
    
    'For 2:
    For i = LBound(matrix) + 1 To UBound(matrix)
        If (Mid(theWord, i, 1) = Mid(theWord, i + 1, 1)) Then
            maxLength = 2
            startAt = i
            tblMatrix.Cells(i, i + 1).Interior.Color = vbYellow
            tblMatrix.Cells(i, i + 1) = 2
        Else
            tblMatrix.Cells(i, i + 1) = 1
        End If
    Next
    
    'For >2:
    Dim k As Long
    For k = 3 To length
        Dim startingIndex As Long
        For startingIndex = 1 To length - k + 1
            Dim endingIndex As Long: endingIndex = startingIndex + k - 1
            
                With tblMatrix
                    .Cells(length + 3, 1) = Mid(theWord, startingIndex, 1)
                    .Cells(length + 3, 2) = Mid(theWord, endingIndex, 1)
                    .Cells(length + 3, 1).Interior.Color = vbRed
                    .Cells(length + 3, 2).Interior.Color = vbRed
                End With
                
                Dim myCell As Range
                Set myCell = tblMatrix.Cells(startingIndex, endingIndex)
                myCell.Select
                                              
                If Mid(theWord, startingIndex, 1) = Mid(theWord, endingIndex, 1) Then
                    myCell.Interior.Color = vbYellow
                    myCell = myCell.Offset(1, -1) + 2
                    maxLength = k
                    startAt = startingIndex
                Else
                    myCell = WorksheetFunction.Max(myCell.Offset(0, -1), myCell.Offset(1, 0))
                End If
        Next startingIndex
    Next k
    
    With tblMatrix
        .Range(.Cells(length + 2, startAt), .Cells(length + 2, startAt + maxLength - 1)).Interior.Color = vbYellow
    End With
    
End Sub

Sub EditTheTable(length As Long)
    
    tblMatrix.Cells.Delete
    Dim i As Long
    For i = 1 To length
        tblMatrix.Columns(i).ColumnWidth = 3.14
    Next
    
End Sub

Sub ClearTable()
    tblMatrix.Cells.Clear
End Sub

Sub WriteTheWord(theWord As String)
    
    Dim row As Long
    Dim col As Long
    Dim sizeCounter As Long
    
    For row = 1 To Len(theWord) + 2
        If row <> Len(theWord) + 1 Then
        For col = 1 To Len(theWord)
            sizeCounter = sizeCounter + 1
            tblMatrix.Cells(row, col) = Mid(theWord, sizeCounter, 1)
        Next
        End If
        sizeCounter = 0
    Next
        
End Sub

