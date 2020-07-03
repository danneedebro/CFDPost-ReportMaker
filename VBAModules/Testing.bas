Attribute VB_Name = "Testing"
Option Explicit

Private Sub TestReplaceMultiple()

    Debug.Print ReplaceMultiple("${1} ${2}", "${1}", "Text1", "${2}", "Text2")
    
    Debug.Print ReplaceMultiple("${1} ${2}", Array("${1}", "Text1", "${2}", "Text2"))
    
    Debug.Print ReplaceMultiple("${1} ${2}", Split("${1};Text1;${2};Text2", ";"))
    
    
    Dim APA As Variant
    APA = Array()
    Debug.Print ReplaceMultiple("${1} ${2}", APA)
    
    Dim stringArray() As String
    stringArray = Split("${1};Text1;${2};Text2", ";")
    Debug.Print ReplaceMultiple("${1} ${2}", stringArray)
    
    Dim variantArray As Variant
    variantArray = Split("${1};Text1;${2};Text2", ";")
    Debug.Print ReplaceMultiple("${1} ${2}", variantArray)
End Sub


Private Sub TestGetRange()
    'On Error GoTo ErrorHandler
    Dim testRange As Range
    Set testRange = GetRange("Template.Plane_")
    Debug.Print testRange.Text
    Exit Sub
ErrorHandler:
    If Err.Number = vbObjectError + 513 Then
        MsgBox "Hejsan"
    Else
        MsgBox "Other error"
    End If
    
End Sub


Private Sub TestSelectTables()
    Dim selectedRange As Range
    Dim table As Variant
    
    For Each table In Array("Views.Names", "UserLocations", "Figures.Geometry", "Figures.Mesh", "Figures.Solution", "ExternalFigures.Convergence", "ExternalFigures.Misc")
        Set selectedRange = GetRange(CStr(table))
        selectedRange.Select
        MsgBox table
    Next table
End Sub



