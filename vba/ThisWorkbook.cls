VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    On Error GoTo OpenError

    Application.ScreenUpdating = False

    ProtectOperationalSheets

    ThisWorkbook.Worksheets("Control_Panel").Activate

CleanExit:

    Application.ScreenUpdating = True
    Exit Sub

OpenError:

    MsgBox "Workbook opened with a protection setup issue: " & _
           Err.Description, _
           vbExclamation, _
           "Consolidation Engine"

    Resume CleanExit

End Sub
