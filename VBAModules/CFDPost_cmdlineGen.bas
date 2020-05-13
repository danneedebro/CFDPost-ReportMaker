Attribute VB_Name = "CFDPost_cmdlineGen"
Option Explicit

Const TABLE_FIGURES = "Figures"
Const TABLE_VIEWS = "Views"
Const FIGURE_HEIGHT = "Figure.Height"
Const FIGURE_WIDTH = "Figure.Width"
Const REPORT_PATH = "Report.Path"
Const REPORT_TITLE = "Report.Title"
Const REPORT_AUTHOR = "Report.Author"

Private Sub Text2Clipboard(OutputString As String)
'
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    MsgBox "Text copied" & vbNewLine & vbNewLine & OutputString, vbOKOnly + vbInformation, "Command line arguments copied"
    
    With MSForms_DataObject
        .SetText OutputString
        .PutInClipboard
    End With

End Sub

Function GetRow(TableName As String, ColumnNum As Long, Key As Variant) As Range
    On Error Resume Next
    Set GetRow = Range(TableName) _
        .Rows(WorksheetFunction.Match(Key, Range(TableName).Columns(ColumnNum), 0))
    If Err.Number <> 0 Then
        Err.Clear
        Set GetRow = Nothing
    End If
End Function

Function NewLines(NumberOfNewLines As Integer) As String
    NewLines = String(NumberOfNewLines, vbNewLine)
End Function

Function Cells2List(ParamArray SelectedCells() As Variant) As String
    Dim retStr As String
    retStr = ""
    Dim c As Variant
    For Each c In SelectedCells
        retStr = IIf(retStr = "", c, retStr & "," & c.Value)
    Next c
    Cells2List = retStr
End Function

Function ReplaceMultiple(TextInput As String, ParamArray Args() As Variant) As String
' Action: Replaces multiple values
'
    If (UBound(Args) - LBound(Args) + 1) Mod 2 <> 0 Then
        MsgBox "Error: arguments not a multiple of 2", vbCritical, "Not a multiple of 2"
        Exit Function
    End If
    
    Dim i As Integer, textOutput As String
    textOutput = TextInput
    For i = 0 To UBound(Args) - 1
        textOutput = Replace(textOutput, Args(i), Args(i + 1))
        i = i + 1
    Next i
    
    ReplaceMultiple = textOutput
    
End Function

Sub ColorWildcards()
' Action: Colors all wildcards (words indicated by ${SOME_TEXT} )
'
    Dim text0 As String, pos1 As Long, pos2 As Long, flag As Boolean
    text0 = ActiveCell.Text
    
    pos1 = 1
    Do
        pos1 = InStr(pos1, text0, "${")
        If pos1 = 0 Then
            flag = False
        Else
            pos2 = InStr(pos1, text0, "}")
            If pos2 = 0 Then
                flag = False
            Else
                Debug.Print "From pos=" & CStr(pos1) & " to pos=" & CStr(pos2)
                flag = True
                ActiveCell.Characters(Start:=pos1, Length:=pos2 - pos1 + 1).Font.Color = RGB(255, 0, 0)
                pos1 = pos2
            End If
        End If
    Loop While flag = True

End Sub


Sub TTest()

    Dim r As Range
    Set r = GetRow("UserLocations", 1, "Plane XY")
    If Not r Is Nothing Then
        r.Select
        MsgBox r.Cells(1, 2).Value
    End If
End Sub




Sub FiguresSetViews()
' Action: Produces an out
'
    Dim Figures As Range, UserLocations As Range
    Set Figures = Range(TABLE_FIGURES)
    Set UserLocations = Range("UserLocations")
    
    Dim i As Integer, j As Integer
    Dim outStr As String
    outStr = ""
    
    ' Define HideAllKnownObjects subroutine
    Dim locType As String
    outStr = outStr & "# Hides all known objects in Figure X" & vbNewLine
    outStr = outStr & "!sub HideAllKnownObjects {" & vbNewLine
    For i = 1 To UserLocations.Rows.Count
        If UserLocations(i, 1) <> "" Then
            outStr = outStr & "   >hide /" & UCase(UserLocations(i, 2)) & ":" & UserLocations(i, 1) & ", view=/VIEW:$_[0]" & vbNewLine
        End If
    Next i
    outStr = outStr & "!}" & vbNewLine
    
    Dim sourceView As String, targetFigure As String, userLocType As String, userLocName As String
    Dim userLocs As Variant
    For i = 1 To Figures.Rows.Count
        sourceView = Figures(i, 2)
        targetFigure = Figures(i, 1)
        outStr = outStr & "# Figure: " & targetFigure & vbNewLine
        outStr = outStr & "!HideAllKnownObjects(""" & targetFigure & """);" & vbNewLine
        userLocs = Split(Figures(i, 4), ",")
        
        For j = LBound(userLocs) To UBound(userLocs)
            userLocName = userLocs(j)
            Dim r As Range
            Set r = GetRow("UserLocations", 1, userLocName)
            If Not r Is Nothing Then
                userLocType = r.Cells(1, 2)
            Else
                Dim question
                question = MsgBox("Error in """ & targetFigure & """, user location """ & userLocName & """ not found", vbExclamation + vbYesNoCancel, "User location not found")
                If question <> vbYes Then Exit Sub Else userLocType = "Plane"
            End If
            outStr = outStr & ">show /" & UCase(userLocType) & ":" & userLocName & ", view=/VIEW:" & targetFigure & vbNewLine
        Next j
        
        
        outStr = outStr & ">delete /VIEW:" & targetFigure & "/CAMERA" & vbNewLine
        outStr = outStr & ">copy from = /VIEW:" & sourceView & "/CAMERA, to = /VIEW:" & targetFigure & "/CAMERA" & vbNewLine
        If Figures(i, 3) = "Yes" Then
            outStr = outStr & "> report showItem=/VIEW:" & targetFigure & vbNewLine
        Else
            outStr = outStr & "> report hideItem=/VIEW:" & targetFigure & vbNewLine
        End If
        
        outStr = outStr & "" & vbNewLine
        
    Next i

    Text2Clipboard outStr

End Sub



Sub MainPreferences()
' Action: Set main pref
'
    Dim outStr As String
    outStr = ">setPreferences Viewer Background Colour Type = Solid, Viewer Background Image File =  , Viewer \" & vbNewLine & "Background Colour = 1.0&1.0&1.0"
        
    Text2Clipboard outStr

End Sub



Sub ReportPublish()
' Action: Set report settings and publish reports
'
    Dim outStr As String, tmpStr As String
    Dim i As Integer
    Dim figureHeight As Range, figureWidth As Range, reportPath As Range
    Set figureHeight = Range(FIGURE_HEIGHT)
    Set figureWidth = Range(FIGURE_WIDTH)
    Set reportPath = Range(REPORT_PATH)
    outStr = ""
    
    outStr = outStr & ReplaceMultiple(Range("TemplateReportSettings").Text, "${FIGURE_HEIGHT}", figureHeight, "${FIGURE_WIDTH}", _
                                      figureWidth, "${REPORT_PATH}", reportPath, "${REPORT_AUTHOR}", Range("Report.Author").Text, _
                                      "${REPORT_TITLE}", Range("Report.Title").Text)

    outStr = outStr & "" & vbNewLine
    
    ' Model description
    tmpStr = "<p><b>Solver:</b><br>" & Range("Solver.Type").Value & ", " & Range("Solver.Time").Value & "</p>"
    tmpStr = tmpStr & "<p><b>Turbulence: " & "</b><br>Model = " & Range("TurbulenceModel.Name") & "<br>Wall function = " & Range("TurbulenceModel.WallFunction").Text & "</p>"
    tmpStr = tmpStr & "<p><b>Fluid: " & Range("Fluid.Description") & "</b><br>Density = " & Range("Fluid.Density") & " kg/m3<br>Viscosity = " & Range("Fluid.Viscosity").Text & " Pa·s</p>"
    tmpStr = tmpStr & ReplaceMultiple(Range("TemplateCommentSubheading").Text, "${TITLE}", "Inlet:", "${TEXT}", Range("BC.Inlet").Text)
    tmpStr = tmpStr & ReplaceMultiple(Range("TemplateCommentSubheading").Text, "${TITLE}", "Outlet:", "${TEXT}", Range("BC.Outlet").Text) & vbNewLine
    outStr = outStr & ReplaceMultiple(Range("TemplateComment").Text, "${COMMENT_NAME}", "Header Description", "${COMMENT_HEADING_LEVEL}", "2", "${COMMENT_HEADING}", "Model description", "${COMMENT_TEXT}", tmpStr)
    outStr = outStr & NewLines(1)
    
    ' External figures - convergence plots
    Dim explotsConvergence As Range
    Set explotsConvergence = Range("ExternalFigures.Convergence")
    tmpStr = ""
    For i = 1 To explotsConvergence.Rows.Count
        If explotsConvergence(i, 3) <> "" Then
            tmpStr = tmpStr & ReplaceMultiple(Range("TemplateExternalFigure").Text, "${PATH}", explotsConvergence(i, 3).Text, Chr(10), "\" & Chr(10), _
                                              "${FIGURE_NAME}", "Figure C" & CStr(i), "${FIGURE_CAPTION}", explotsConvergence(i, 2).Text)
        End If
    Next i
    outStr = outStr & ReplaceMultiple(Range("TemplateComment").Text, "${COMMENT_NAME}", "Header convergence", "${COMMENT_HEADING_LEVEL}", "2", "${COMMENT_HEADING}", "Convergence", "${COMMENT_TEXT}", tmpStr)
    outStr = outStr & "" & vbNewLine
    
    ' External figures - Misc
    Dim explotsMisc As Range
    Set explotsMisc = Range("ExternalFigures.Misc")
    tmpStr = ""
    For i = 1 To explotsMisc.Rows.Count
        If explotsMisc(i, 3) <> "" Then
            tmpStr = tmpStr & ReplaceMultiple(Range("TemplateExternalFigure").Text, "${PATH}", explotsMisc(i, 3).Text, Chr(10), "\" & Chr(10), _
                                              "${FIGURE_NAME}", "Figure M" & CStr(i), "${FIGURE_CAPTION}", explotsMisc(i, 2).Text)
        End If
    Next i
    outStr = outStr & ReplaceMultiple(Range("TemplateComment").Text, "${COMMENT_NAME}", "Header Misc", "${COMMENT_HEADING_LEVEL}", "2", "${COMMENT_HEADING}", "Misc", "${COMMENT_TEXT}", tmpStr)
    outStr = outStr & "" & vbNewLine
    
    
    outStr = outStr & "> update" & NewLines(2)
    outStr = outStr & ">report save" & vbNewLine
    
    Text2Clipboard outStr
    
End Sub

