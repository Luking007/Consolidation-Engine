Attribute VB_Name = "mod_Protection"
Option Explicit

Public Sub ProtectOperationalSheets()

    ProtectOneSheet ThisWorkbook.Worksheets("Master_Data")
    ProtectOneSheet ThisWorkbook.Worksheets("Import_Log")
    ProtectOneSheet ThisWorkbook.Worksheets("Data_Quality_Log")

End Sub


Private Sub ProtectOneSheet(ByVal ws As Worksheet)

    ' Reapply protection every time the workbook opens.
    ' UserInterfaceOnly:=True allows VBA to write to protected
    ' sheets while preventing normal manual edits by the user.

    ws.Unprotect Password:=""

    ws.Protect _
        Password:="", _
        UserInterfaceOnly:=True, _
        AllowFiltering:=True

End Sub

Public Sub UnprotectOperationalSheets()

    ThisWorkbook.Worksheets("Master_Data").Unprotect Password:=""
    ThisWorkbook.Worksheets("Import_Log").Unprotect Password:=""
    ThisWorkbook.Worksheets("Data_Quality_Log").Unprotect Password:=""

End Sub
