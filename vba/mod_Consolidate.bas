Attribute VB_Name = "mod_Consolidate"
Sub AppendToMaster(rowsToAppend As Collection, sourceFile As String, batchID As String)
    If rowsToAppend.Count = 0 Then Exit Sub

    Dim tbl As ListObject
    Set tbl = Sheets("Master_Data").ListObjects(1)

    Dim outputArray() As Variant
    ReDim outputArray(1 To rowsToAppend.Count, 1 To 13)

    Dim r As Long, c As Long, oneRow As Variant
    Dim stampTime As Date: stampTime = Now

    For r = 1 To rowsToAppend.Count
        oneRow = rowsToAppend(r)
        For c = 1 To 10
            outputArray(r, c) = oneRow(c)
        Next c
        outputArray(r, 11) = sourceFile
        outputArray(r, 12) = batchID
        outputArray(r, 13) = stampTime
    Next r

    Dim i As Long, newListRow As ListRow, firstNewListRow As ListRow
    For i = 1 To rowsToAppend.Count
        Set newListRow = tbl.ListRows.Add
        If i = 1 Then Set firstNewListRow = newListRow
    Next i

    firstNewListRow.Range.Resize(rowsToAppend.Count, 13).value = outputArray
End Sub

Sub TestAppendToMaster()
    Dim los As Variant
    Dim sourceProfiles As Object
    Dim testRows As New Collection

    los = ImportWorkbook(ThisWorkbook.Path & "\Input\LOS_Station_Report.xlsx")
    Set sourceProfiles = LoadSourceProfiles()

    ' Real RemapRow output this time, not a hand-typed Array(...) --
    ' genuinely 1-indexed, exactly what this function actually
    ' receives once ProcessOneFile exists to call it for real.
    testRows.Add RemapRow(los, 2, sourceProfiles("Standard"))  ' Aug 1
    testRows.Add RemapRow(los, 3, sourceProfiles("Standard"))  ' Aug 2

    Dim tbl As ListObject
    Set tbl = Sheets("Master_Data").ListObjects(1)
    Debug.Print "Rows before: "; tbl.ListRows.Count

    AppendToMaster testRows, "TEST_AppendToMaster.xlsx", "TEST-BATCH"

    Debug.Print "Rows after: "; tbl.ListRows.Count
    Debug.Print "First new row, Station_Code: "; tbl.DataBodyRange.Cells(tbl.ListRows.Count - 1, 1).value
    Debug.Print "Second new row, Flights_Handled: "; tbl.DataBodyRange.Cells(tbl.ListRows.Count, 4).value
End Sub
