Attribute VB_Name = "mod_Main"
Option Explicit

Sub RunFullImport()
    ' The one button-click entry point. Discovers every .xlsx in
    ' /Input/, runs each through the full pipeline (import -> map ->
    ' validate -> dedupe -> consolidate -> log), and reports a summary.
    ' One bad FILE can't crash the run -- ProcessOneFile already has
    ' its own local error trap, proven when it was built. This handler
    ' only catches genuinely batch-breaking setup failures.
    On Error GoTo FatalError

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    UnprotectOperationalSheets

    Dim profiles As Object, rules As Object, seenKeys As Object
    Set profiles = LoadSourceProfiles()
    Set rules = LoadValidationRules()
    Set seenKeys = CreateObject("Scripting.Dictionary")

    ' One Batch_ID for the WHOLE run, shared across every file -- ties
    ' LOS, ABV, PHC and KAN together as "the same import" in the logs,
    ' and lets seenKeys catch a duplicate whether it repeats within one
    ' file or turns up again in a later one, same run.
    Dim batchID As String
    batchID = "BATCH-" & Format(Now, "yyyymmdd-hhnnss")

    Dim files As Collection
    Set files = GetFilesInFolder(ThisWorkbook.Path & "\Input\", "*.xlsx")

    If files.Count = 0 Then
        LogEvent Sheets("Import_Log"), "RUN " & batchID & ": no .xlsx files found in /Input/", batchID:=batchID, status:="WARNING"
        
        ProtectOperationalSheets
        
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        MsgBox "No files found in the Input folder.", vbExclamation, "Nothing to Import"
        Exit Sub
    End If

    Dim f As Variant, okCount As Long, failCount As Long
    For Each f In files
        If ProcessOneFile(CStr(f), profiles, rules, seenKeys, batchID) Then
            okCount = okCount + 1
        Else
            failCount = failCount + 1
        End If
    Next f

    LogEvent Sheets("Import_Log"), "RUN " & batchID & " COMPLETE: " & okCount & " OK, " & failCount & " failed", batchID:=batchID, status:="COMPLETE"

ProtectOperationalSheets

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.DisplayAlerts = True

    Dim summary As String
    summary = "Import complete." & vbCrLf & vbCrLf & _
        "Files processed OK: " & okCount & vbCrLf & _
        "Files failed: " & failCount & vbCrLf & vbCrLf & _
        "Master_Data now has " & Sheets("Master_Data").ListObjects(1).ListRows.Count & " rows." & vbCrLf & _
        "Check Data_Quality_Log for anything flagged."
    MsgBox summary, vbInformation, "Consolidation Complete"
    Exit Sub

FatalError:

    On Error Resume Next

    ProtectOperationalSheets

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    LogEvent Sheets("Import_Log"), "FATAL: " & Err.Description

    MsgBox "Import halted before it could start: " & Err.Description, _
           vbCritical, _
           "Fatal Error"

End Sub
