Attribute VB_Name = "mod_Import"
Option Explicit

Function ImportWorkbook(filePath As String) As Variant
    ' Opens filePath read-only, reads the whole used range (header row
    ' + data) into a 2D array, closes without saving, hands back the
    ' array. No error handling in here on purpose -- ProcessOneFile
    ' will wrap this call in its own local trap later, same nested-
    ' handler pattern as mod_Main. This function just does one job.
    Dim wb As Workbook
    Dim rawData As Variant

    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    rawData = wb.Worksheets(1).UsedRange.value
    wb.Close SaveChanges:=False

    ImportWorkbook = rawData
End Function

Function ProcessOneFile(filePath As String, profiles As Object, rules As Object, seenKeys As Object, batchID As String) As Boolean
    On Error GoTo FileError

    Dim fileName As String
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    Dim raw As Variant
    raw = ImportWorkbook(filePath)

    Dim profileName As String
    profileName = MatchProfile(raw, profiles)
    If profileName = "" Then
        LogEvent Sheets("Import_Log"), "UNMATCHED FORMAT: " & fileName, batchID, fileName, "FAILED"
        ProcessOneFile = False
        Exit Function
    End If

    ' RemapRow handles ONE data row per call -- loop every data row
    ' (row 1 = header) to build the full set before validation starts.
    Dim mappedRows As New Collection
    Dim r As Long
    For r = 2 To UBound(raw, 1)
        mappedRows.Add RemapRow(raw, r, profiles(profileName))
    Next r

    Dim cleanRows As Collection, flaggedRows As Collection
    SplitCleanAndFlagged mappedRows, rules, cleanRows, flaggedRows
    LogFlaggedRows flaggedRows, fileName, batchID

    Dim deduped As Collection
    Set deduped = RemoveDuplicates(cleanRows, seenKeys)

    AppendToMaster deduped, fileName, batchID

    LogEvent Sheets("Import_Log"), "OK: " & fileName & " -- " & mappedRows.Count & " rows, " & _
    flaggedRows.Count & " flagged, " & (cleanRows.Count - deduped.Count) & " duplicate, " & deduped.Count & " written", _
    batchID, fileName, "OK", deduped.Count, flaggedRows.Count, (cleanRows.Count - deduped.Count)

    ProcessOneFile = True
    Exit Function

FileError:
    LogEvent Sheets("Import_Log"), "FILE FAILED: " & fileName & " -- " & Err.Description, batchID, fileName, "FAILED"
    ProcessOneFile = False
End Function

