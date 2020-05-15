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

Sub CreateReportSkeleton()
' Action: Creates the skeleton for the state file (.cst)
'
    Dim outStr As New ClassString, i As Integer
    
    ' STEP 0: Load result file
    outStr.AppendRow "!sub LoadResultFile{"
    outStr.AppendRow ">close"
    outStr.AppendRow OutputLoadFile()
    outStr.AppendRow "!}"
    
    ' STEP 1: Create user locations and plots
    outStr.AppendRow "!sub CreateUserLocationsAndPlots{"
    outStr.AppendRow OutputUserLocationsAndPlots()
    outStr.AppendRow "!}"
    
    ' STEP 2: Create model description
    outStr.AppendRow "!sub UpdateModelDescription{"
    outStr.AppendRow OutputModelDescription()
    outStr.AppendRow "!}"
    
    ' STEP 3: Create Result table
    outStr.AppendRow "!sub UpdateResultTable{"
    outStr.AppendRow OutputResultTable()
    outStr.AppendRow "!}"
    
    ' STEP 4: Create figures and report elements
    Dim figureTables As Variant, figureTable As Variant, figureTableRng As Range
    figureTables = Array("Geometry", "Mesh", "Solution")
    outStr.AppendRow "!sub CreateFigures{"
    For Each figureTable In figureTables
        outStr.AppendRow ReplaceMultiple(Range("Template.Comment"), "${NAME}", "Header " & figureTable, "${COMMENT_HEADING_LEVEL}", "1", _
                                         "${COMMENT_HEADING}", figureTable, "${COMMENT_TEXT}", "")
        Set figureTableRng = Range("Figures." & figureTable)
        
        For i = 1 To figureTableRng.Rows.Count
            If figureTableRng(i, 1).Text <> "" Then
                outStr.AppendRow ReplaceMultiple(">delete /VIEW:${FIGURE}", "${FIGURE}", figureTableRng(i, 1))
                outStr.AppendRow ReplaceMultiple("> setViewportView cmd=shallow_copy, view=/VIEW:${FIGURE}, viewport=1", "${FIGURE}", figureTableRng(i, 1))
            End If
        Next i
    Next figureTable
    outStr.AppendRow "!}"
    
    ' STEP 5: Reset views
    outStr.AppendRow "!sub ResetViews{"
    outStr.AppendRow OutputResetViews()
    outStr.AppendRow "!}"
    
    ' STEP 5: Create external plots
    outStr.AppendRow "!sub LoadExternalFigures{"
    outStr.AppendRow OutputExternalFigures()
    outStr.AppendRow "!}"
    
    ' STEP 6: Update report settings
    outStr.AppendRow "!sub SetReportSettings{"
    outStr.AppendRow OutputReportSettings()
    outStr.AppendRow "!}"
    
    ' STEP 7: Set report order
    outStr.AppendRow "!sub SortReportItems{"
    outStr.AppendRow OutputReportOrder()
    outStr.AppendRow "!}"
    
    ' STEP 8: Publish report
    outStr.AppendRow "!sub PublishReport{"
    outStr.AppendRow "REPORT:"
    outStr.AppendRow "PUBLISH:"
    outStr.AppendRow "    Report Path = $_[0]"
    outStr.AppendRow "  END"
    outStr.AppendRow "END"
    outStr.AppendRow "> update"
    outStr.AppendRow ">report save"
    outStr.AppendRow "!}"
    
    outStr.NewLines 2
    outStr.AppendRow "# Comment out the subroutines not to run"
    outStr.AppendRow "# Step 1. Load result and create"
    outStr.AppendRow "!LoadResultFile();"
    outStr.AppendRow "!CreateUserLocationsAndPlots();"
    outStr.AppendRow "!CreateFigures();"
    outStr.AppendRow "# Step 2. Manually adjust User locations and plots"
    outStr.AppendRow "# Step 3. Manually adjust the camera of View 1 to View 4"
    outStr.AppendRow "# Step 4. Comment out subs in Step 1 and run subs below"
    outStr.AppendRow "!ResetViews();"
    outStr.AppendRow "!UpdateModelDescription();"
    outStr.AppendRow "!UpdateResultTable();"
    outStr.AppendRow "!LoadExternalFigures();"
    outStr.AppendRow "!SortReportItems();"
    outStr.AppendRow "!SetReportSettings();"
    outStr.AppendRow "# Step 5: Publish report"
    outStr.AppendRow ReplaceMultiple("# !PublishReport(""${PATH}"");", "${PATH}", Range("Report.Path"))

    Text2Clipboard outStr.Output
    
End Sub

Function OutputModelDescription() As String
' Action: Returns text for the model description
'
    Dim tmpStr As String, outStr As String
    
    tmpStr = "<p><b>Solver:</b><br>" & Range("Solver.Type").Value & ", " & Range("Solver.Time").Value & "</p>"
    tmpStr = tmpStr & "<p><b>Turbulence: " & "</b><br>Model = " & Range("TurbulenceModel.Name") & "<br>Wall function = " & Range("TurbulenceModel.WallFunction").Text & "</p>"
    tmpStr = tmpStr & "<p><b>Fluid: " & Range("Fluid.Description") & "</b><br>Density = " & Range("Fluid.Density") & " kg/m3<br>Viscosity = " & Range("Fluid.Viscosity").Text & " Pa·s</p>"
    tmpStr = tmpStr & ReplaceMultiple(Range("Template.CommentSubheading").Text, "${TITLE}", "Inlet:", "${TEXT}", Range("BC.Inlet").Text)
    tmpStr = tmpStr & ReplaceMultiple(Range("Template.CommentSubheading").Text, "${TITLE}", "Outlet:", "${TEXT}", Range("BC.Outlet").Text) & vbNewLine
    OutputModelDescription = ReplaceMultiple(Range("Template.Comment").Text, "${NAME}", "Header Description", "${COMMENT_HEADING_LEVEL}", "1", "${COMMENT_HEADING}", "Model description", "${COMMENT_TEXT}", tmpStr)
End Function

Function OutputResultTable() As String
' Action: Returns the result table
'
    Dim outStr As New ClassString, i As Integer, tableInput As Range
    outStr.AppendRow "TABLE:Result Table"
    outStr.AppendRow "  TABLE CELLS:"
    Set tableInput = Range("TableInput")
    For i = 1 To tableInput.Rows.Count
        If tableInput(i, 1) <> "" Then
            outStr.AppendRow ReplaceMultiple("    ${CELL} = ""${FORMULA}"", False, False, False, Left, True, 0, Font Name, 1|1, %10.3e, True, ffffff, 000000, True", _
                                             "${CELL}", tableInput(i, 1), "${FORMULA}", tableInput(i, 2))
        End If
    Next i
        
    outStr.AppendRow "  END"
    outStr.AppendRow "END"
    
    OutputResultTable = outStr.Output

End Function

Function OutputReportOrder() As String
' Action: Returns cmdstr for setting the report order
'
    Dim outStr As New ClassString
    outStr.Append "REPORT:" & vbNewLine & "  Report Items = /TITLE PAGE,/REPORT/FILE INFORMATION OPTIONS,/REPORT/MESH STATISTICS OPTIONS," & _
                  "/REPORT/PHYSICS SUMMARY OPTIONS,/REPORT/SOLUTION SUMMARY OPTIONS,/REPORT/OPERATING MAPS," & _
                  "/COMMENT:Header Description,/COMMENT:User Data,/TABLE:Result Table"
    
    ' Loop through figure tables (Geometry, Mesh and Solution)
    Dim figureTables As Variant, figureTable As Variant, figureTableRng As Range, i As Integer
    figureTables = Array("Geometry", "Mesh", "Solution")
    
    For Each figureTable In figureTables
        Set figureTableRng = Range("Figures." & figureTable)
        outStr.Append ",/COMMENT:Header " & figureTable
        
        For i = 1 To figureTableRng.Rows.Count
            If figureTableRng(i, 1).Text <> "" Then
                outStr.Append ",/VIEW:" & figureTableRng(i, 1)
            End If
        Next i
    Next figureTable
    
    outStr.Append ",/COMMENT:Header Convergence,/COMMENT:Header Misc" & vbNewLine & "END"
    
    OutputReportOrder = outStr.Output
    
End Function

Function OutputExternalFigures() As String
' Action: Returns text for external figures
'
    ' External figures - convergence plots
    Dim figureTables As Variant, figureTable As Variant, figureTableRng As Range, i As Integer
    Dim tmpStr As String, outStr As String
    figureTables = Array("Convergence", "Misc")
    
    For Each figureTable In figureTables
        Set figureTableRng = Range("ExternalFigures." & figureTable)
        
        tmpStr = ""
        For i = 1 To figureTableRng.Rows.Count
            If figureTableRng(i, 3) <> "" Then
                tmpStr = tmpStr & ReplaceMultiple(Range("Template.ExternalFigure").Text, "${PATH}", figureTableRng(i, 3).Text, Chr(10), "\" & Chr(10), _
                                                  "${FIGURE_NAME}", "Figure C" & CStr(i), "${FIGURE_CAPTION}", figureTableRng(i, 2).Text)
            End If
        Next i
        outStr = outStr & ReplaceMultiple(Range("Template.Comment").Text, "${NAME}", "Header " & figureTable, "${COMMENT_HEADING_LEVEL}", "1", "${COMMENT_HEADING}", figureTable, "${COMMENT_TEXT}", tmpStr)
        outStr = outStr & "" & vbNewLine
    Next figureTable
    
    OutputExternalFigures = outStr
End Function

Function OutputUserLocationsAndPlots() As String
' Action: Returns the cmdstr to create the user locations (Planes, surface groups, etc) and plots (Contour, Streamline, etc)
'
    Dim outStr As New ClassString
    Dim objectTypes As Variant, objectType As Variant
    objectTypes = Array("Wireframe", "Plane", "Surface group", "Contour", "Streamline")
    
    Dim i As Integer
    Dim userLocations As Range
    Set userLocations = Range("UserLocations")
    Dim firstUserLocation As String
    firstUserLocation = "inlet"
    
    For Each objectType In objectTypes
        For i = 1 To userLocations.Rows.Count
            If userLocations(i, 2) = objectType Then
                Select Case objectType
                    Case "Wireframe"
                        outStr.AppendRow ReplaceMultiple(Range("Template.Wireframe"), "${NAME}", userLocations(i, 1))
                    Case "Plane"
                        If firstUserLocation = "inlet" Then firstUserLocation = userLocations(i, 1)
                        outStr.AppendRow PlanePointAndNormal(userLocations(i, 1), 0, 0, 0, 0, 0, 1)
                        outStr.NewLines 1
                    Case "Surface group"
                        outStr.AppendRow ReplaceMultiple(Range("Template.SurfaceGroup"), "${NAME}", userLocations(i, 1), "${LOCATION_LIST}", firstUserLocation, "${TRANSPARENCY}", "0.2")
                        outStr.NewLines 1
                    Case "Contour"
                        outStr.AppendRow ReplaceMultiple(Range("Template.Contour"), "${NAME}", userLocations(i, 1), "${VARIABLE}", "Pressure", "${LOCATION_LIST}", firstUserLocation)
                        outStr.NewLines 1
                    Case "Streamline"
                        outStr.AppendRow ReplaceMultiple(Range("Template.Streamline"), "${NAME}", userLocations(i, 1))
                        outStr.NewLines 1
                End Select
            End If
        Next i
    Next objectType
    
    OutputUserLocationsAndPlots = outStr.Output
    
End Function

Function OutputResetViews() As String
' Action: Returns the cmdstr to reset the view (activate all objects in all figures)
'
    Dim outStr As New ClassString
    
    Dim i As Integer, j As Integer
    'Dim outStr As String
    'outStr = ""
    
    ' Define HideAllKnownObjects subroutine
    Dim userLocations As Range
    Set userLocations = Range("UserLocations")
    Dim locType As String
    outStr.AppendRow "# Hides all known objects in Figure X"
    outStr.AppendRow "!sub HideAllKnownObjects {"
    For i = 1 To userLocations.Rows.Count
        If userLocations(i, 1) <> "" Then
            outStr.AppendRow "   >hide /" & UCase(userLocations(i, 2)) & ":" & userLocations(i, 1) & ", view=/VIEW:$_[0]"
        End If
    Next i
    outStr.AppendRow "!}"
    
    Dim figures As Range, figureTables As Variant, figureTable As Variant
    Dim sourceView As String, targetFigure As String, userLocType As String, userLocName As String
    Dim userLocs As Variant
    figureTables = Array("Figures.Geometry", "Figures.Mesh", "Figures.Solution")
    
    ' Loop through each figureTable (Geometry, Mesh and Solution) and hide all objects then
    ' show the ones that suppose to be shown
    For Each figureTable In figureTables
        Set figures = Range(figureTable)
        For i = 1 To figures.Rows.Count
            sourceView = figures(i, 2)
            targetFigure = figures(i, 1)
            outStr.AppendRow "# Figure: " & targetFigure
            outStr.AppendRow "!HideAllKnownObjects(""" & targetFigure & """);"
            userLocs = Split(figures(i, 4), ",")
            
            For j = LBound(userLocs) To UBound(userLocs)
                userLocName = userLocs(j)
                Dim r As Range
                Set r = GetRow("UserLocations", 1, userLocName)
                If Not r Is Nothing Then
                    userLocType = r.Cells(1, 2)
                Else
                    Dim question
                    question = MsgBox("Error in """ & targetFigure & """, user location """ & userLocName & """ not found", vbExclamation + vbYesNoCancel, "User location not found")
                    If question <> vbYes Then Exit Function Else userLocType = "Plane"
                End If
                outStr.AppendRow ">show /" & UCase(userLocType) & ":" & userLocName & ", view=/VIEW:" & targetFigure
            Next j
            
            
            outStr.AppendRow ">delete /VIEW:" & targetFigure & "/CAMERA"
            outStr.AppendRow ">copy from = /VIEW:" & sourceView & "/CAMERA, to = /VIEW:" & targetFigure & "/CAMERA"
            If figures(i, 3) = "Yes" Then
                outStr.AppendRow "> report showItem=/VIEW:" & targetFigure
            Else
                outStr.AppendRow "> report hideItem=/VIEW:" & targetFigure
            End If
            
            outStr.AppendRow ""
            
        Next i
    Next figureTable
    
    OutputResetViews = outStr.Output
    
End Function

Function OutputReportSettings() As String
' Action: Returns cmdstr for setting the report settings
'
    Dim outStr As New ClassString

    ' Update report settings
    outStr.AppendRow ReplaceMultiple(Range("Template.ReportSettings").Text, "${FIGURE_HEIGHT}", Range("Figure.Height"), "${FIGURE_WIDTH}", _
                                      Range("Figure.Width"), "${REPORT_PATH}", Range("Report.Path"), "${REPORT_AUTHOR}", Range("Report.Author").Text, _
                                      "${REPORT_TITLE}", Range("Report.Title").Text)
    ' White background
    outStr.AppendRow ">setPreferences Viewer Background Colour Type = Solid, Viewer Background Image File =  , Viewer Background Colour = 1.0&1.0&1.0"
    
    OutputReportSettings = outStr.Output
    
End Function

Function OutputLoadFile() As String
' Action: Returns the cmdstr for loading a file
'
    Dim outStr As New ClassString
    outStr.AppendRow ReplaceMultiple(Range("Template.LoadFile"), "${FILENAME}", Range("ResultFile"))
    
    OutputLoadFile = outStr.Output
    
End Function

Function PlanePointAndNormal(Name As String, X0 As Double, Y0 As Double, Z0 As Double, NormalX As Double, NormalY As Double, NormalZ As Double) As String
' Action:
'
    Dim pointStr As String, normalVectorStr As String
    pointStr = Replace("X0 [m], Y0 [m], Z0 [m]", "X0", CStr(X0))
    pointStr = Replace(pointStr, "Y0", CStr(Y0))
    pointStr = Replace(pointStr, "Z0", CStr(Z0))
    normalVectorStr = Replace("NormX , NormY , NormZ", "NormX", CStr(NormalX))
    normalVectorStr = Replace(normalVectorStr, "NormY", CStr(NormalY))
    normalVectorStr = Replace(normalVectorStr, "NormZ", CStr(NormalZ))
    
    PlanePointAndNormal = ReplaceMultiple(Range("Template.Plane"), "${NAME}", Name, "${NORMAL}", normalVectorStr, "${POINT}", pointStr)

End Function

