Attribute VB_Name = "mod_Dedupe"
Option Explicit

Function RemoveDuplicates(cleanRows As Collection, ByVal seenKeys As Object) As Collection
    ' Seed seenKeys with records already present in Master_Data.
    ' This makes rerunning the same source files safe: existing
    ' Station_Code + Report_Date combinations will not be appended again.
    If seenKeys.Count = 0 Then
        SeedExistingMasterKeys seenKeys
    End If

    Dim survivors As New Collection
    Dim row As Variant
    Dim key As String

    For Each row In cleanRows
        key = BuildDedupeKey(row)

        If Not seenKeys.Exists(key) Then
            seenKeys.Add key, True
            survivors.Add row
        End If

        ' Existing keys are treated as duplicates and are not added
        ' to survivors. Import_Log records the duplicate count for
        ' the current source file.
    Next row

    Set RemoveDuplicates = survivors
End Function


Private Sub SeedExistingMasterKeys(ByVal seenKeys As Object)

    Dim tbl As ListObject
    Set tbl = Sheets("Master_Data").ListObjects(1)

    ' Nothing to seed when Master_Data is empty.
    If tbl.ListRows.Count = 0 Then Exit Sub

    Dim r As Long
    Dim stationCode As String
    Dim reportDate As Variant
    Dim key As String

    For r = 1 To tbl.ListRows.Count

        stationCode = CStr(tbl.DataBodyRange.Cells(r, 1).value)
        reportDate = tbl.DataBodyRange.Cells(r, 3).value

        ' Only build a key when the required fields exist.
        If Len(stationCode) > 0 And IsDate(reportDate) Then

            key = stationCode & "|" & _
                  Format(CDate(reportDate), "yyyy-mm-dd")

            If Not seenKeys.Exists(key) Then
                seenKeys.Add key, True
            End If

        End If

    Next r

End Sub


Private Function BuildDedupeKey(row As Variant) As String

    ' Station_Code + Report_Date define the business key.
    Const STATION_CODE = 1
    Const REPORT_DATE = 3

    BuildDedupeKey = CStr(row(STATION_CODE)) & "|" & _
                     Format(row(REPORT_DATE), "yyyy-mm-dd")

End Function

