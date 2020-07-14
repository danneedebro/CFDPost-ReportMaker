Attribute VB_Name = "HelperFunctions"
Option Explicit
Option Private Module

Function FindRowInRange(TableName As String, ColumnNum As Long, Key As Variant) As Range
' Action: Return a row (as a range object) from a named range (TableName)
'
    On Error Resume Next
    Set FindRowInRange = GetRange(TableName) _
        .Rows(WorksheetFunction.Match(Key, Range(TableName).Columns(ColumnNum), 0))
    If Err.Number <> 0 Then
        Err.Clear
        Set FindRowInRange = Nothing
    End If
End Function

Function GetRange(RangeName As String, Optional RaiseError As Boolean = True) As Range
' Action: Return a named range
'
    On Error GoTo ErrorHandler
    ThisWorkbook.Activate
    Set GetRange = Range(RangeName)
    Exit Function
ErrorHandler:
    If RaiseError = True Then
        Err.Raise vbObjectError + 513, "Function GetRange", "Named range """ & RangeName & """ not found. ActiveWorkbook=" & ActiveWorkbook.Name & ", ActiveSheet=" & ActiveSheet.Name
    Else
        Set GetRange = Nothing
    End If
End Function

Function Indent(StringBlockToIndent As String, Optional NumberOfSpaces As Integer = 4) As String
' Action: Indents the supplied string with
'
    Indent = Replace(String(NumberOfSpaces, " ") & StringBlockToIndent, Chr(10), Chr(10) & String(NumberOfSpaces, " "))
End Function

Function NewLines(NumberOfNewLines As Integer) As String
    NewLines = String(NumberOfNewLines, vbNewLine)
End Function


Function ReplaceMultiple(TextInput As String, ParamArray Arguments() As Variant) As String
' Action: Replaces multiple values
'
    ' Handle cases where variable arguments both given as Array("","") and comma-separated values
    Dim args As Variant
    Dim nArgs As Integer
    nArgs = UBound(Arguments) - LBound(Arguments) + 1
    If nArgs = 0 Then
        args = Arguments
        ReplaceMultiple = TextInput
        Exit Function
    ElseIf nArgs = 1 And VarType(Arguments(UBound(Arguments))) >= vbArray Then
        args = Arguments(0)
    Else
        args = Arguments
    End If

    If (UBound(args) - LBound(args) + 1) Mod 2 <> 0 Then
        Err.Raise vbObjectError + 514, "ReplaceMultiple", "Error: arguments not a multiple of 2" & vbNewLine & Join(args, ";")
        'MsgBox "Error: arguments not a multiple of 2", vbCritical, "Not a multiple of 2"
        Exit Function
    End If
    
    Dim i As Integer, textOutput As String
    textOutput = TextInput
    For i = 0 To UBound(args) - 1
        textOutput = Replace(textOutput, args(i), args(i + 1))
        i = i + 1
    Next i
    
    ReplaceMultiple = textOutput
    
End Function
