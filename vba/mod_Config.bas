Attribute VB_Name = "mod_Config"
Option Explicit

Function LoadSourceProfiles() As Object
    ' Reads tbl_SourceProfiles into memory once per run.
    ' Returns a Dictionary of Dictionaries:
    '   profiles("Standard")("MatchPattern")        -> "Station_Code"
    '   profiles("Standard")("Fuel_Uplift_Litres")  -> Array("Fuel_Uplift_Litres", False)
    '   profiles("ABV_Style")("Station_Code")       -> Array("Airport Code", True)
    ' i.e. one inner dictionary per profile, holding one entry per
    ' field it maps, plus a special "MatchPattern" entry.
    Dim profiles As Object
    Set profiles = CreateObject("Scripting.Dictionary")

    Dim tbl As ListObject
    Set tbl = Sheets("Config_SourceProfiles").ListObjects(1)

    Dim data As Variant
    data = tbl.DataBodyRange.value   ' whole table, one read -- no cell-by-cell loop

    Dim i As Long
    Dim profileName As String, standardField As String
    Dim sourceHeader As String, matchPattern As String, requiredText As String
    Dim innerDict As Object

    For i = 1 To UBound(data, 1)
        profileName = data(i, 1)    ' column A: Profile_Name
        matchPattern = data(i, 2)   ' column B: Match_Pattern
        standardField = data(i, 3)  ' column C: Standard_Field
        sourceHeader = data(i, 4)   ' column D: Source_Header
        requiredText = data(i, 5)   ' column E: Required

        If Not profiles.Exists(profileName) Then
            Set innerDict = CreateObject("Scripting.Dictionary")
            innerDict.Add "MatchPattern", matchPattern
            profiles.Add profileName, innerDict
        End If

        Set innerDict = profiles.Item(profileName)
        innerDict.Add standardField, Array(sourceHeader, (requiredText = "Yes"))
    Next i

    Set LoadSourceProfiles = profiles
End Function

Function LoadValidationRules() As Object
    ' Reads tbl_ValidationRules into memory once per run.
    ' Returns a Dictionary keyed by Field name; each value is itself
    ' a Dictionary holding that field's rule details:
    '   rules("Station_Code")("AllowedValues") -> Array("LOS","ABV","PHC","KAN")
    '   rules("Report_Date")("MinValue")       -> a real Date
    ' One rule per field, so this is one level of nesting, not two
    ' like LoadSourceProfiles.
    Dim rules As Object
    Set rules = CreateObject("Scripting.Dictionary")

    Dim tbl As ListObject
    Set tbl = Sheets("Config_ValidationRules").ListObjects(1)

    Dim data As Variant
    data = tbl.DataBodyRange.value   ' whole table, one read

    Dim i As Long
    Dim fieldName As String
    Dim ruleDict As Object

    For i = 1 To UBound(data, 1)
        fieldName = data(i, 1)   ' column A: Field

        Set ruleDict = CreateObject("Scripting.Dictionary")
        ruleDict.Add "RuleType", data(i, 2)      ' column B
        ruleDict.Add "MinValue", data(i, 3)      ' column C
        ruleDict.Add "MaxValue", data(i, 4)      ' column D

        If Len(data(i, 5)) > 0 Then               ' column E: Allowed_Values
            ruleDict.Add "AllowedValues", Split(data(i, 5), ",")
        Else
            ruleDict.Add "AllowedValues", Array()
        End If

        ruleDict.Add "ErrorMessage", data(i, 6)  ' column F

        rules.Add fieldName, ruleDict
    Next i

    Set LoadValidationRules = rules
End Function
