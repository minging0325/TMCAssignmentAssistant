Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim wsAssignments As Worksheet
    Dim wsPathwayLog As Worksheet
    Dim clickedCell As Range
    Dim assignmentName As String
    Dim sortRange As Range

    ' Set the Assignments worksheet
    Set wsAssignments = ThisWorkbook.Sheets("Assignments")
    Set wsPathwayLog = ThisWorkbook.Sheets("Pathway Log")
    ' Check if a single cell is selected
    If Target.Cells.Count <> 1 Or Target.Row < 5 Or Target.Row > 18 Then
        Exit Sub
    End If
    
    Set clickedCell = Target
    If clickedCell Is Nothing Or Len(clickedCell.Value) <> 0 Then
        Exit Sub
    End If
    
    ' Get the content of column A (Assignment name)
    assignmentName = wsAssignments.Cells(clickedCell.Row, 1).Value
    ' MsgBox "Assignment name: " & assignmentName  ' Print Assignment name immediately

    ' Determine the sorting range
    Set sortRange = FindSortRange(wsAssignments, assignmentName, clickedCell.Column)
    ' MsgBox "Sorting range: " & sortRange.Address
    
    ' Get memberList
    memberNames = Application.WorksheetFunction.Transpose(FindMemberListRange(wsPathwayLog))
    
    ' Get sorted memberList
    sortedMemberNames = SortNameListByFrequency(sortRange, memberNames)
    sortedMemberNames = RemoveNameFromLastMeetingWithSameRole(sortRange, sortedMemberNames)
    sortedMemberNames = RemoveNameFromTheSameMeeting(wsAssignments, sortedMemberNames, clickedCell.Column)
    For i = LBound(sortedMemberNames) To UBound(sortedMemberNames)
       ' Debug.Print sortedMemberNames(i)
    Next i
    
    ' Write the sorted names into the dropdown menu
    AddDropDown sortedMemberNames, clickedCell
End Sub

Function FindMemberListRange(wsPathwayLog As Worksheet) As Range
    Dim lastRow As Long
    Dim rowWithEmptyCell As Long
    Dim cell As Range
    
    ' Find the last non-empty cell in column A of wsPathwayLog worksheet
    lastRow = wsPathwayLog.Cells(wsPathwayLog.Rows.Count, "A").End(xlUp).Row
    
    ' Start searching from A2 downwards to find the first empty cell
    For rowWithEmptyCell = 2 To lastRow
        If wsPathwayLog.Cells(rowWithEmptyCell, "A").Value = "" Then
            ' If an empty cell is found, exit the loop
            Exit For
        End If
    Next rowWithEmptyCell
    
    ' If an empty cell exists, set the range from A2 to the cell above the empty cell
    If rowWithEmptyCell > 2 Then
        Set FindMemberListRange = wsPathwayLog.Range("A2:A" & rowWithEmptyCell - 1)
    Else
        ' If no empty cell exists, set the range from A2 to the last non-empty cell
        Set FindMemberListRange = wsPathwayLog.Range("A2:A" & lastRow)
    End If
End Function

Function FindSortRange(ws As Worksheet, assignmentName As String, clickedCol As Long) As Range
    Dim lastRow As Long
    Dim i As Long
    Dim startCol As Long
    Dim endCol As Long

    Dim startRow As Long
    Dim endRow As Long
    startRow = -1
    endRow = -1
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ' Find matching rows
    For i = 1 To lastRow
        If ws.Cells(i, 1).Value = assignmentName Then
            If startRow = -1 Then
                startRow = i
                endRow = i
            Else
                endRow = i
            End If
        End If
    Next i
    startCol = 2
    ' endCol = ws.Cells(startRow, ws.Columns.Count).End(xlToLeft).Column
    endCol = clickedCol - 1
    ' MsgBox "endCol: " & endCol
    Set FindSortRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(endRow, endCol))
    
End Function

Function RemoveNameFromTheSameMeeting(ws As Worksheet, nameList As Variant, clickedCol As Long) As Variant()
    For rowIdx = 5 To 18
        cellValue = ws.Cells(rowIdx, clickedCol).Value
        If Not IsEmpty(cellValue) Then
            For nameIdx = LBound(nameList) To UBound(nameList)
                ' Check if the name in nameList matches the name in (row, lastCol)
                If nameList(nameIdx) = cellValue Then
                    ' Remove the matching name
                    nameList(nameIdx) = ""
                End If
            Next nameIdx
        End If
    Next rowIdx
    RemoveNameFromTheSameMeeting = nameList
End Function

Function RemoveNameFromLastMeetingWithSameRole(rng As Range, nameList As Variant) As Variant()
    Dim lastCol As Long
    lastCol = rng.Columns.Count
    For rowIdx = 1 To rng.Rows.Count
        ' Get the value of (row, lastCol) cell
        Dim cellValue As Variant
        cellValue = rng.Cells(rowIdx, lastCol).Value
    
        ' Check if the (row, lastCol) cell is not empty
        If Not IsEmpty(cellValue) Then
            ' Iterate through nameList
            For nameIdx = LBound(nameList) To UBound(nameList)
                ' Check if the name in nameList matches the name in (row, lastCol)
                If nameList(nameIdx) = cellValue Then
                    ' Remove the matching name
                    nameList(nameIdx) = ""
                End If
            Next nameIdx
        End If
    Next rowIdx
    RemoveNameFromLastMeetingWithSameRole = nameList
End Function

Function SortNameListByFrequency(rng As Range, nameList As Variant) As Variant()
    Dim nameCounts As Object
    Set nameCounts = CreateObject("Scripting.Dictionary")

    Dim nameCount As Integer
    Dim name As Variant
    Dim sortedList As Variant
    Dim lastColName As Variant
    Dim i As Long
    Dim result() As Variant
    Dim resultIndex As Long

    ' Iterate through the Name List, counting each name's occurrences in the Range
    For Each name In nameList
        If name <> "" Then
            ' Count the occurrences of the name in the rightmost 8 columns of the range
            nameCount = 0
            For colOffset = 0 To 6
                If lastCol - colOffset = 0 Then
                    Exit For ' Exit the loop when lastCol - colOffset = 0
                End If
                nameCount = nameCount + Application.WorksheetFunction.CountIf(rng.Columns(lastCol - colOffset), name)
            Next colOffset

            ' Add the name and its count to the dictionary
           

 nameCounts.Add name, nameCount
        End If
    Next name
    
    ' You can use nameCounts here, for example, print it to the Immediate window
    For Each key In nameCounts.Keys
        Debug.Print key, nameCounts(key)
    Next key
    
    ' Use a sorting subroutine
    Dim sortedNameCounts() As Variant
    sortedNameCounts = SortNameCounts(nameCounts)

    ' You can use sortedNameCounts here, for example, print it to the Immediate window
    For i = LBound(sortedNameCounts) To UBound(sortedNameCounts)
        ' Debug.Print sortedNameCounts(i)
    Next i
    ' Return the sorted list of names
    SortNameListByFrequency = sortedNameCounts
End Function

Sub Swap(ByRef arr As Variant, ByVal i As Long, ByVal j As Long)
    Dim temp As Variant
    temp = arr(i)
    arr(i) = arr(j)
    arr(j) = temp
End Sub

Function SortNameCounts(nameCounts As Object) As Variant()
    ' Store names and counts in an array
    Dim keyValueArray() As Variant
    Dim key As Variant
    Dim i As Integer
    i = 1
    ReDim keyValueArray(1 To nameCounts.Count, 1 To 2)

    For Each key In nameCounts.Keys
        keyValueArray(i, 1) = key
        keyValueArray(i, 2) = nameCounts(key)
        i = i + 1
    Next key

    ' Sort the array in ascending order based on counts
    QuickSort keyValueArray, LBound(keyValueArray, 1), UBound(keyValueArray, 1), 2

    ' Build sortedNameCounts
    Dim sortedNameCounts() As Variant
    ReDim sortedNameCounts(1 To UBound(keyValueArray, 1))

    For i = 1 To UBound(keyValueArray, 1)
        sortedNameCounts(i) = keyValueArray(i, 1)
    Next i

    ' Return the sorted list of names
    SortNameCounts = sortedNameCounts
End Function

Sub QuickSort(arr() As Variant, low As Long, high As Long, sortColumn As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As Variant
    Dim temp As Variant

    i = low
    j = high
    pivot = arr((low + high) \ 2, sortColumn)

    Do While i <= j
        Do While arr(i, sortColumn) < pivot
            i = i + 1
        Loop
        Do While arr(j, sortColumn) > pivot
            j = j - 1
        Loop
        If i <= j Then
            ' Swap
            For Col = LBound(arr, 2) To UBound(arr, 2)
                temp = arr(i, Col)
                arr(i, Col) = arr(j, Col)
                arr(j, Col) = temp
            Next Col
            i = i + 1
            j = j - 1
        End If
    Loop

    If low < j Then QuickSort arr, low, j, sortColumn
    If i < high Then QuickSort arr, i, high, sortColumn
End Sub

Sub AddDropDown(sortedNameCounts As Variant, clickedCell As Range)
    Dim dataValidation As Validation
    
    ' Clear existing data validation
    clickedCell.Validation.Delete
    
    ' Add data validation
    Set dataValidation = clickedCell.Validation
    With dataValidation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=Join(sortedNameCounts, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
End Sub