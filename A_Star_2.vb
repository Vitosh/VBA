Option Explicit

Public C_COLUMNS        As Long
Public C_ROWS           As Long

Public cell_start       As Range
Public cell_end         As Range
Public Sub SetColsAndRows()

    C_COLUMNS = tbl_matrix.tb_cols
    C_ROWS = tbl_matrix.tb_rows

End Sub
Public Sub Main()
    
    Dim cell_current As Range

    Dim l_smallest_path As Long
    Dim l_col As Long
    Dim l_C_ROWS As Long
    
   On Error GoTo Main_Error
    
    Call ObstaclesFromSelect
    Call SetColsAndRows
    Call SetCellStart
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

   On Error GoTo 0
   Exit Sub

Main_Error:
    MsgBox "No Way", vbOKOnly, "No Way"
    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of Module mod_main"
    
End Sub
Public Sub ObstaclesFromSelect()
    
    Dim r_intersect As Object
    
    Set r_intersect = Application.intersect(Selection, [Playground])
    
    If Not tbl_matrix.cb_obstacles Then
    
        If Not r_intersect Is Nothing Then
            r_intersect.Style = "Accent1"
        End If
    End If
    
    Set r_intersect = Nothing
    
End Sub

Public Sub SetCellStart()
    
    Set cell_start = Cells(1, 2)
    Set cell_end = Cells(C_ROWS, C_COLUMNS)

End Sub

Public Sub Reset()
    
    Dim sName           As String
    Dim rCell           As Range
    
    Call SetColsAndRows
    Call SetCellStart
    
    Cells.Clear
    Range(Cells(1, 2), Cells(C_ROWS, C_COLUMNS)).Name = "Playground"
    


    [Playground].Style = "Neutral"
    [Playground].RowHeight = 14
    [Playground].ColumnWidth = 2.3
    [Playground].WrapText = True
    
    Call ObstaclesFromSelect
    Call MakeProblems
    
    cell_start.Style = "Bad"
    cell_end.Style = "Good"

End Sub
Public Sub AdvertiseHere()

    Range(Cells(C_ROWS + 1, 2), Cells(C_ROWS + 1, C_COLUMNS)).Merge
    Range(Cells(C_ROWS + 1, 2), Cells(C_ROWS + 1, C_COLUMNS)) = "Vitoshacademy.com!"
    Range(Cells(C_ROWS + 1, 2), Cells(C_ROWS + 1, C_COLUMNS)).HorizontalAlignment = xlCenter
    
End Sub
Public Sub MakeProblems()
    
    Dim dbl_row             As Double
    Dim dbl_col             As Double
    Dim dbl_counter         As Variant
    Dim r_cell              As Range
    
    dbl_counter = tbl_matrix.tb_obstacles
        While dbl_counter > 0
        
            dbl_row = Int((C_ROWS - 2 + 1) * Rnd + 2)
            dbl_col = Int((C_COLUMNS - 2 + 1) * Rnd + 2)
            If dbl_row + dbl_col <> 3 And dbl_row + dbl_col <> C_ROWS + C_COLUMNS Then
                Set r_cell = Cells(dbl_row, dbl_col)
                r_cell.Style = "Accent3"
            End If
            dbl_counter = dbl_counter - 1
        Wend
        
    Set r_cell = Nothing
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
        my_cell = cell_current.Address & "*" & distance_to_success(my_cell) & "*" & price_to_reach(my_cell, cell_current) & "*" & distance_to_success(my_cell) + price_to_reach(my_cell, cell_current)
    End If
    
End Function

Public Function price_to_reach(ByRef my_cell, ByRef cell_current) As Double
    
    Dim d_diagonal_1                As Double
    Dim d_diagonal_2                As Double
    
    Dim d_diagonal_difference       As Double
    Dim l_straight_difference       As Double
    
    d_diagonal_1 = Abs(my_cell.Row - cell_current.Row)
    d_diagonal_2 = Abs(my_cell.Column - cell_current.Column)
    d_diagonal_difference = Application.WorksheetFunction.Min(d_diagonal_1, d_diagonal_2)
    
    l_straight_difference = Abs(Abs(my_cell.Row - cell_current.Row) + Abs(my_cell.Column - cell_current.Column))
    l_straight_difference = l_straight_difference - 2 * d_diagonal_difference
    price_to_reach = l_straight_difference * 10 + d_diagonal_difference * 14
    
    If Not cell_current = "" Then
        price_to_reach = price_to_reach + Split(cell_current, "*")(2)
    End If
    
End Function

Public Function distance_to_success(my_cell As Range) As Double
    
    Dim d_diagonal_1                As Double
    Dim d_diagonal_2                As Double
    
    Dim d_diagonal_difference       As Double
    Dim l_straight_difference       As Double
    
    d_diagonal_1 = Abs(my_cell.Row - cell_end.Row)
    d_diagonal_2 = Abs(my_cell.Column - cell_end.Column)
    d_diagonal_difference = Application.WorksheetFunction.Min(d_diagonal_1, d_diagonal_2)
    
    l_straight_difference = Abs(Abs(my_cell.Row - cell_end.Row) + Abs(my_cell.Column - cell_end.Column))
    l_straight_difference = l_straight_difference - 2 * d_diagonal_difference
    distance_to_success = l_straight_difference * 10 + d_diagonal_difference * 14

End Function

Public Function find_possible_smallest_path(ByRef current_cell As Range) As Range

    Dim my_cell             As Range
    Dim my_result_cell      As Range
    Dim l_result            As Long
    
    l_result = 1000000000
    Set my_result_cell = Nothing
    
    For Each my_cell In [Playground]
        If my_cell.Style = "Calculation" Then
            If CDbl(Split(my_cell, "*")(1)) + CDbl(Split(my_cell, "*")(2)) < l_result Then
                l_result = CDbl(Split(my_cell, "*")(1)) + CDbl(Split(my_cell, "*")(2))
                Set my_result_cell = my_cell
            End If
        End If
    Next my_cell
    Set find_possible_smallest_path = my_result_cell
    Set my_result_cell = Nothing
    
End Function


