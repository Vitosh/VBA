Option Explicit

Public Const C_COLUMNS = 60
Public Const C_ROWS = 20

Public cell_start       As Range
Public cell_end         As Range

Public Sub Main()
    Dim cell_current As Range

    Dim l_smallest_path As Long
    Dim l_col As Long
    Dim l_C_ROWS As Long
    
    Call Reset
    Set cell_current = cell_start

    Do While True
        If check_for_success(cell_current) Then Exit Do

        Set cell_current = find_possible_smallest_path(cell_current)
        cell_current.Style = "Input"
        
    Loop

    Do While True
        cell_current.Style = "Accent2"
        If check_for_success(cell_current, False) Then Exit Do
        Set cell_current = Range(Split(cell_current, "*")(0))
        
    Loop
    
    Call AdvertiseHere
    
    Set cell_start = Nothing
    Set cell_end = Nothing
    Set cell_current = Nothing
    
End Sub

Public Sub Reset()
    
    Dim sName As String
    
    Cells.Delete
    Range(Cells(1, 1), Cells(C_ROWS, C_COLUMNS)).Name = "Playground"
    
    Set cell_start = Cells(1, 1)
    Set cell_end = Cells(C_ROWS, C_COLUMNS)

    [playground].Style = "Neutral"
    [playground].RowHeight = 14
    [playground].ColumnWidth = 2.3
    [playground].WrapText = True
    
    
    Call MakeProblems
    
    cell_start.Style = "Bad"
    cell_end.Style = "Good"
    
End Sub
Public Sub AdvertiseHere()

    Range(Cells(C_ROWS + 1, 1), Cells(C_ROWS + 1, C_COLUMNS)).Merge
    Range(Cells(C_ROWS + 1, 1), Cells(C_ROWS + 1, C_COLUMNS)) = "Vitoshacademy.com!"
    Range(Cells(C_ROWS + 1, 1), Cells(C_ROWS + 1, C_COLUMNS)).HorizontalAlignment = xlCenter
    
End Sub
Public Sub MakeProblems()
    
    Selection.Style = "Accent1"

End Sub

Public Function check_for_success(ByRef cell_current As Range, Optional b_going_back As Boolean = True) As Boolean
    
    Dim my_cell As Range
    
    '3
    If cell_current.Column < C_COLUMNS Then
        Set my_cell = cell_current.Offset(0, 1)
        check_for_success = ChangeCellData(my_cell, b_going_back, cell_current)
        If check_for_success Then Exit Function
    End If
    
    '4.5
    If cell_current.Column < C_COLUMNS And cell_current.Row < C_ROWS Then
        Set my_cell = cell_current.Offset(1, 1)
        check_for_success = ChangeCellData(my_cell, b_going_back, cell_current)
        If check_for_success Then Exit Function
    End If
    
    '6
    If cell_current.Row < C_ROWS Then
        Set my_cell = cell_current.Offset(1, 0)
        check_for_success = ChangeCellData(my_cell, b_going_back, cell_current)
        If check_for_success Then Exit Function
    End If
    
    '7.5
    If cell_current.Column > 1 And cell_current.Row < C_ROWS Then
        Set my_cell = cell_current.Offset(1, -1)
        check_for_success = ChangeCellData(my_cell, b_going_back, cell_current)
        If check_for_success Then Exit Function
    End If
        
    '9
    If cell_current.Column > 1 Then
        Set my_cell = cell_current.Offset(0, -1)
        check_for_success = ChangeCellData(my_cell, b_going_back, cell_current)
        If check_for_success Then Exit Function
    End If
    
    '11.5
    If cell_current.Column > 1 And cell_current.Row > 1 Then
        Set my_cell = cell_current.Offset(-1, -1)
        check_for_success = ChangeCellData(my_cell, b_going_back, cell_current)
        If check_for_success Then Exit Function
    End If
    
    '12
    If cell_current.Row > 1 Then
        Set my_cell = cell_current.Offset(-1, 0)
        check_for_success = ChangeCellData(my_cell, b_going_back, cell_current)
        If check_for_success Then Exit Function
    End If
    
    '1.5
    If cell_current.Column < C_COLUMNS And cell_current.Row > 1 Then
        Set my_cell = cell_current.Offset(-1, 1)
        check_for_success = ChangeCellData(my_cell, b_going_back, cell_current)
        If check_for_success Then Exit Function
    End If
    
    Set my_cell = Nothing
    
End Function

Public Function ChangeCellData(ByRef my_cell As Range, ByRef b_going_back As Boolean, cell_current As Range) As Boolean
    
    If my_cell.Style = IIf(b_going_back, "Good", "Bad") Then ChangeCellData = True
    
    If my_cell.Style = "Neutral" Then
        my_cell.Style = "Calculation"
        my_cell = cell_current.Address & "*" & distance_to_success(my_cell)
    End If
    
End Function

Public Sub ColorNeighbours(cell_current)

End Sub

Public Function distance_to_success(my_cell As Range) As Long
    
    distance_to_success = Abs(my_cell.Row - cell_end.Row) + Abs(my_cell.Column - cell_end.Column)
    
End Function

Public Function find_possible_smallest_path(ByRef current_cell As Range) As Range

    Dim my_cell             As Range
    Dim my_result_cell      As Range
    Dim l_result            As Long
    
    l_result = 1000000000
    Set my_result_cell = Nothing
    
    For Each my_cell In [playground]
        If my_cell.Style = "Calculation" Then
            If Split(my_cell, "*")(1) < l_result Then
                l_result = Split(my_cell, "*")(1)
                Set my_result_cell = my_cell
            End If
        End If
    Next my_cell
    
    Set find_possible_smallest_path = my_result_cell
    Set my_result_cell = Nothing
    
End Function
