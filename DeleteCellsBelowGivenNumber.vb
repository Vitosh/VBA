sub DeleteCells
	Dim lCellsToBeDeleted As Long
	lCellsToBeDeleted = Cells(Rows.Count, "A").End(xlUp).Row
	If lCellsToBeDeleted > 9 Then
	  Rows("10:" & CStr(lCellsToBeDeleted)).Delete Shift:=xlUp
	End If
end sub