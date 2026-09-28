Attribute VB_Name = "mod_Mapping"
Option Explicit

Function MatchProfile(rawData As Variant, profiles As Object) As String
    ' Looks at rawData's header row (row 1) and returns the name of
    ' the first profile whose MatchPattern text appears in any header
    ' cell. Returns "" if nothing matches -- that's not an error here,
    ' it's a real outcome ProcessOneFile will need to handle later
    ' (log it, skip the file, move on).
    Dim profileName As Variant
    Dim pattern As String
    Dim col As Long

    For Each profileName In profiles.Keys
        pattern = profiles(profileName)("MatchPattern")

        For col = 1 To UBound(rawData, 2)
            If InStr(1, CStr(rawData(1, col)), pattern, vbTextCompare) > 0 Then
                MatchProfile = CStr(profileName)
                Exit Function
            End If
        Next col
    Next profileName

    MatchProfile = ""
End Function

Function RemapRow(rawData As Variant, dataRowIndex As Long, ByVal profile As Object) As Variant
    ' rawData is a full imported table -- row 1 is always the header
    ' row, dataRowIndex picks which data row (2, 3, 4...) to pull
    ' values from. Returns that row re-ordered into the standard field
    ' sequence, using profile's field mappings.
    Dim standardFields As Variant
    standardFields = Array("Station_Code", "Station_Name", "Report_Date", _
        "Flights_Handled", "OnTime_Departures", "Delayed_Departures", _
        "Cancelled_Flights", "Avg_Turnaround_Min", "Fuel_Uplift_Litres", _
        "Passengers_Handled")

    Dim output(1 To 10) As Variant
    Dim i As Long, col As Long, sourceHeader As String, foundCol As Long

    For i = 0 To UBound(standardFields)
        If profile.Exists(standardFields(i)) Then
            sourceHeader = profile(standardFields(i))(0)

            foundCol = 0
            For col = 1 To UBound(rawData, 2)
                If CStr(rawData(1, col)) = sourceHeader Then
                    foundCol = col
                    Exit For
                End If
            Next col

            If foundCol > 0 Then
                If standardFields(i) = "Report_Date" Then
    ' Every source routes through SafeParseDate here -- so whatever
    ' comes OUT of RemapRow is always a real Date, regardless of
    ' whether the file stored it as a real date or PHC's text format.
    ' Nothing downstream (validation, dedupe, consolidation) ever
    ' has to know or care which one a given file used.
    output(i + 1) = SafeParseDate(rawData(dataRowIndex, foundCol))
Else
    output(i + 1) = rawData(dataRowIndex, foundCol)
End If
            Else
                output(i + 1) = Null   ' e.g. KAN has no Fuel_Uplift_Litres column
            End If
        Else
            output(i + 1) = Null
        End If
    Next i

    RemapRow = output
End Function

