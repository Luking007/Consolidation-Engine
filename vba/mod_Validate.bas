Attribute VB_Name = "mod_Validate"
Option Explicit

Function ValidateRow(remappedRow As Variant, rules As Object) As String
    ' remappedRow is RemapRow's output -- 10 values in standard field
    ' order. Checks every simple per-field rule from Config_ValidationRules,
    ' plus one hardcoded cross-field rule that table can't express.
    ' Returns "" if the row is clean, or a semicolon-separated list of
    ' every problem found -- never just the first one.
    Const STATION_CODE = 1, REPORT_DATE = 3, FLIGHTS_HANDLED = 4
    Const ONTIME = 5, DELAYED = 6, CANCELLED = 7
    Const TURNAROUND = 8, FUEL = 9, PASSENGERS = 10

    Dim problems As String
    problems = ""

    ' --- Station_Code: AllowedList ---
    If Not IsInAllowedList(remappedRow(STATION_CODE), rules("Station_Code")("AllowedValues")) Then
        problems = problems & rules("Station_Code")("ErrorMessage") & "; "
    End If

    ' --- Report_Date: DateRange ---
    If IsDate(remappedRow(REPORT_DATE)) Then
        If remappedRow(REPORT_DATE) < rules("Report_Date")("MinValue") _
        Or remappedRow(REPORT_DATE) > rules("Report_Date")("MaxValue") Then
            problems = problems & rules("Report_Date")("ErrorMessage") & "; "
        End If
    Else
        problems = problems & "Report_Date could not be read as a date; "
    End If

    ' --- Every plain numeric Range field ---
    problems = problems & CheckRange(remappedRow(FLIGHTS_HANDLED), rules("Flights_Handled"))
    problems = problems & CheckRange(remappedRow(ONTIME), rules("OnTime_Departures"))
    problems = problems & CheckRange(remappedRow(DELAYED), rules("Delayed_Departures"))
    problems = problems & CheckRange(remappedRow(CANCELLED), rules("Cancelled_Flights"))
    problems = problems & CheckRange(remappedRow(TURNAROUND), rules("Avg_Turnaround_Min"))
    problems = problems & CheckRange(remappedRow(PASSENGERS), rules("Passengers_Handled"))

    ' --- Fuel_Uplift_Litres: only checked when present (KAN has none) ---
    If Not IsNull(remappedRow(FUEL)) Then
        problems = problems & CheckRange(remappedRow(FUEL), rules("Fuel_Uplift_Litres"))
    End If

    ' --- Hardcoded cross-field rule: not in Config_ValidationRules,
    ' because it compares three fields to each other, not one field
    ' to a fixed bound -- exactly the design decision from Step 2. ---
    If (remappedRow(ONTIME) + remappedRow(DELAYED) + remappedRow(CANCELLED)) > remappedRow(FLIGHTS_HANDLED) Then
        problems = problems & "On-Time + Delayed + Cancelled exceeds Flights Handled; "
    End If

    ValidateRow = problems
End Function

Private Function IsInAllowedList(value As Variant, allowedValues As Variant) As Boolean
    Dim v As Variant
    For Each v In allowedValues
        If CStr(value) = CStr(v) Then
            IsInAllowedList = True
            Exit Function
        End If
    Next v
    IsInAllowedList = False
End Function

Private Function CheckRange(value As Variant, rule As Object) As String
    If value < rule("MinValue") Or value > rule("MaxValue") Then
        CheckRange = rule("ErrorMessage") & "; "
    Else
        CheckRange = ""
    End If
End Function

Sub SplitCleanAndFlagged(mappedRows As Collection, rules As Object, ByRef cleanRows As Collection, ByRef flaggedRows As Collection)
    Set cleanRows = New Collection
    Set flaggedRows = New Collection

    Dim row As Variant, problems As String
    For Each row In mappedRows
        problems = ValidateRow(row, rules)
        If problems = "" Then
            cleanRows.Add row
        Else
            flaggedRows.Add Array(row, problems)
        End If
    Next row
End Sub
