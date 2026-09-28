Attribute VB_Name = "mod_Utilities"
Option Explicit

Function SafeParseDate(rawValue As Variant) As Variant
    ' Returns a proper Date, or Null if rawValue can't be understood.
    ' Never raises an error -- a bad date gets flagged like any other
    ' invalid row later, it doesn't crash the whole file.
    On Error GoTo ParseFailed

    ' Case 1: already a real date (LOS, KAN, ABV all wrote real dates).
    If VarType(rawValue) = vbDate Then
        SafeParseDate = rawValue
        Exit Function
    End If

    ' Case 2: text in DD-MM-YYYY, PHC's format -- parsed explicitly,
    ' never handed to CDate. CDate guesses the format from Windows'
    ' regional settings, and on a PC set to MM/DD, "01-08-2026" would
    ' silently become January 8th instead of August 1st. Wrong date,
    ' no error raised -- worse than a failure, since nothing would
    ' ever flag it.
    If VarType(rawValue) = vbString Then
        Dim parts() As String
        parts = Split(rawValue, "-")
        If UBound(parts) = 2 Then
            SafeParseDate = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
            Exit Function
        End If
    End If

ParseFailed:
    SafeParseDate = Null
End Function

Sub LogEvent(logSheet As Worksheet, msg As String, _
             Optional batchID As String = "", _
             Optional fileName As String = "", _
             Optional status As String = "", _
             Optional rowsImported As Variant, _
             Optional rowsFlagged As Variant, _
             Optional rowsDuplicate As Variant)
    Dim tbl As ListObject
    Set tbl = logSheet.ListObjects(1)

    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add

    newRow.Range(tbl.ListColumns("Timestamp").Index).value = Now
    newRow.Range(tbl.ListColumns("Message").Index).value = msg

    If batchID <> "" Then newRow.Range(tbl.ListColumns("Batch_ID").Index).value = batchID
    If fileName <> "" Then newRow.Range(tbl.ListColumns("File_Name").Index).value = fileName
    If status <> "" Then newRow.Range(tbl.ListColumns("Status").Index).value = status
    If Not IsMissing(rowsImported) Then newRow.Range(tbl.ListColumns("Rows_Imported").Index).value = rowsImported
    If Not IsMissing(rowsFlagged) Then newRow.Range(tbl.ListColumns("Rows_Flagged").Index).value = rowsFlagged
    If Not IsMissing(rowsDuplicate) Then newRow.Range(tbl.ListColumns("Rows_Duplicate").Index).value = rowsDuplicate
End Sub

Function GetFilesInFolder(folderPath As String, pattern As String) As Collection
    ' Returns full paths of every file in folderPath matching pattern,
    ' e.g. GetFilesInFolder(ThisWorkbook.Path & "\Input\", "*.xlsx")
    Dim result As New Collection
    Dim fullFolderPath As String
    Dim fileName As String

    ' Guarantee exactly one backslash, whether or not the caller
    ' included a trailing one -- ThisWorkbook.Path never has one.
    fullFolderPath = folderPath
    If Right(fullFolderPath, 1) <> "\" Then
        fullFolderPath = fullFolderPath & "\"
    End If

    fileName = Dir(fullFolderPath & pattern)
    Do While fileName <> ""
        ' Skip Excel's own hidden lock file (~$File.xlsx), created the
        ' instant a matching workbook is open anywhere -- including,
        ' right now, on your machine, since you've had these sample
        ' files open to read headers off them. Dir() can't tell a lock
        ' file from a real one; only the name gives it away.
        If Left(fileName, 2) <> "~$" Then
            result.Add fullFolderPath & fileName
        End If
        fileName = Dir()  ' no argument -- continues the SAME search
    Loop

    Set GetFilesInFolder = result
End Function

Sub LogFlaggedRows(flaggedRows As Collection, sourceFile As String, batchID As String)
    ' flaggedRows holds Array(remappedRow, problemsText) pairs from
    ' SplitCleanAndFlagged. One bulk write to Data_Quality_Log, same
    ' principle as AppendToMaster.
    If flaggedRows.Count = 0 Then Exit Sub

    Dim tbl As ListObject
    Set tbl = Sheets("Data_Quality_Log").ListObjects(1)

    Dim outputArray() As Variant
    ReDim outputArray(1 To flaggedRows.Count, 1 To 6)

    Dim i As Long, pair As Variant, oneRow As Variant, reason As String
    For i = 1 To flaggedRows.Count
        pair = flaggedRows(i)      ' Array(...) is 0-indexed -- pair(0), pair(1)
        oneRow = pair(0)
        reason = pair(1)

        outputArray(i, 1) = batchID
        outputArray(i, 2) = Now
        outputArray(i, 3) = sourceFile
        outputArray(i, 4) = "Row " & i
        outputArray(i, 5) = reason
        outputArray(i, 6) = JoinRow(oneRow)
    Next i

    Dim j As Long, newListRow As ListRow, firstNewListRow As ListRow
    For j = 1 To flaggedRows.Count
        Set newListRow = tbl.ListRows.Add
        If j = 1 Then Set firstNewListRow = newListRow
    Next j

    firstNewListRow.Range.Resize(flaggedRows.Count, 6).value = outputArray
End Sub

Private Function JoinRow(oneRow As Variant) As String
    ' oneRow is RemapRow's output -- genuinely 1-indexed (Dim output(1 To 10)).
    Dim parts(1 To 10) As String
    Dim k As Long
    For k = 1 To 10
        If IsNull(oneRow(k)) Then
            parts(k) = "(none)"
        Else
            parts(k) = CStr(oneRow(k))   ' CStr(Empty) = "" safely; CStr(Null) would error, hence the guard above
        End If
    Next k
    JoinRow = Join(parts, " | ")
End Function
