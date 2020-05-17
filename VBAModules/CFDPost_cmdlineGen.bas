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
' Action: Return a row (as a range object) from a named range (TableName)
'
    On Error Resume Next
    Set GetRow = Range(TableName) _
        .Rows(WorksheetFunction.Match(Key, Range(TableName).Columns(ColumnNum), 0))
    If Err.Number <> 0 Then
        Err.Clear
        Set GetRow = Nothing
    End If
End Function

Function GetRange(RangeName As String) As Range
' Action: Return a named range
'
    On Error GoTo ErrorHandler
    Set GetRange = Range(RangeName)
    Exit Function
ErrorHandler:
    Err.Raise vbObjectError + 513, "Function GetRange", "Named range """ & RangeName & """ not found. ActiveWorkbook=" & ActiveWorkbook.Name & ", ActiveSheet=" & ActiveSheet.Name
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

Function ArgList(ParamArray Arguments() As Variant) As String
    ' Handle cases where variable arguments both given as Array("","") and comma-separated values
    Dim args As Variant
    Dim nArgs As Integer
    nArgs = UBound(Arguments) - LBound(Arguments) + 1
    If nArgs = 0 Then
        args = Arguments
        
        Exit Function
    ElseIf nArgs = 1 And VarType(Arguments(UBound(Arguments))) >= vbArray Then
        args = Arguments(0)
    Else
        args = Arguments
    End If

    If (UBound(args) - LBound(args) + 1) Mod 2 <> 0 Then
        ArgList = "Error: arguments not a multiple of 2"
        Exit Function
    End If
    
    Dim i As Integer, outStr As String
    For i = LBound(args) To UBound(args)
        outStr = outStr + CStr(args(i)) & IIf(i < UBound(args), ";", "")
    Next i
    ArgList = outStr
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

Function GetWildcards(templateString As String) As Collection
' Action: Returns a collection with all wildcards (words indicated by ${SOME_TEXT} )
'
    Dim pos1 As Long, pos2 As Long, flag As Boolean
    Dim wildcardsFound As New Collection
    
    pos1 = 1
    Do
        pos1 = InStr(pos1, templateString, "${")
        If pos1 = 0 Then
            flag = False
        Else
            pos2 = InStr(pos1, templateString, "}")
            If pos2 = 0 Then
                flag = False
            Else
                wildcardsFound.Add Mid(templateString, pos1, pos2 - pos1 + 1)
                flag = True
                pos1 = pos2
            End If
        End If
    Loop While flag = True
    
    Set GetWildcards = wildcardsFound

End Function

Public Sub SetArguments()
' Action: Sets arguments
'
    Dim userLocations As Range, userLocation As Range
    Set userLocations = GetRange("UserLocations")
    
    ' Get selected row
    With userLocations
        If ActiveCell.Row >= .Row And ActiveCell.Row <= .Row + .Rows.Count - 1 And ActiveCell.Column >= .Column And ActiveCell.Column <= .Column + .Columns.Count - 1 Then
            Debug.Print "Inside"
            Set userLocation = userLocations.Rows(ActiveCell.Row - userLocations.Row + 1)
            
        Else
            Debug.Print "Outside"
            MsgBox "Select a cell inside UserLocations"
            Exit Sub
        End If
    End With
    
    ' Get user location default template and arguments
    Dim objectType As String, userLocationDefault As Range
    Dim defaultArgs As Variant, defaultTemplateName As String
    objectType = userLocation.Cells(1, 2)
    Set userLocationDefault = GetRow("UserLocationDefaults", 1, objectType)
    If Not userLocationDefault Is Nothing Then
        defaultTemplateName = userLocationDefault.Cells(1, 2)
        defaultArgs = Split(userLocationDefault.Cells(1, 3), ";")
    Else
        Dim question
        question = MsgBox("Default user location not found", vbExclamation + vbYesNoCancel, "User location not found")
        If question <> vbYes Then Exit Sub
    End If
    
    ' Replace
    Dim templateName As String, templateString As String
    If userLocation.Cells(1, 3) <> "" Then
        templateName = userLocation.Cells(1, 3)
    Else
        templateName = defaultTemplateName
    End If
    templateString = GetRange(templateName).Text
    
    ' Dim replace
    Dim wildcards As Collection, wildcard As Variant
    Set wildcards = GetWildcards(templateString)
    
    Dim inps As New Collection
    Dim inp As Variant, i As Integer, outStr As String
    For Each wildcard In wildcards
        If wildcard <> "${NAME}" Then
            
            inp = Application.InputBox(wildcard, Type:=0)
            If inp <> False Then
                inps.Add """" & wildcard & """"
                inps.Add Right(inp, Len(inp) - 1)
            End If
        End If
    Next wildcard
    
    For i = 1 To inps.Count
        outStr = outStr & inps(i) & IIf(i = inps.Count, "", ",")
    Next i
    
    userLocation.Cells(1, 4).Formula = "=ArgList(" & outStr & ")"
        
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

Function OutputUserLocationsAndPlots_old() As String
' Action: Returns the cmdstr to create the user locations (Planes, surface groups, etc) and plots (Contour, Streamline, etc)
'
    Dim outStr As New ClassString
    Dim objectTypes As Variant, objectType As Variant
    objectTypes = Array("Wireframe", "Plane", "Surface group", "Contour", "Streamline")
    
    Dim i As Integer
    Dim userLocations As Range
    Set userLocations = GetRange("UserLocations")
    Dim firstUserLocation As String, templateString As String, userArgs As Variant
    firstUserLocation = "inlet"
    
    On Error GoTo ErrorHandler
    For Each objectType In objectTypes
        userLocationDefault
    
    
        For i = 1 To userLocations.Rows.Count
            templateString = ""
            If userLocations(i, 2) = objectType Then
                ' If template given
                If userLocations(i, 3) <> "" Then
                    templateString = GetRange(userLocations(i, 3))
                End If
                If userLocations(i, 4) <> "" Then
                    userArgs = Split(userLocations(i, 4), ";")
                Else
                    userArgs = Array()
                End If
                
                Select Case objectType
                    Case "Wireframe"
                        If templateString = "" Then templateString = GetRange("Template.Wireframe") ' Load default if no template given
                        templateString = ReplaceMultiple(templateString, userArgs)
                        outStr.AppendRow ReplaceMultiple(templateString, "${NAME}", userLocations(i, 1))
                    Case "Plane"
                        If firstUserLocation = "inlet" Then firstUserLocation = userLocations(i, 1)
                        If templateString = "" Then templateString = GetRange("Template.Plane") ' Load default if no template given
                        templateString = ReplaceMultiple(templateString, userArgs)
                        outStr.AppendRow ReplaceMultiple(templateString, "${NAME}", userLocations(i, 1), "${NORMAL}", "0 , 0 , 1", "${POINT}", "0 [m], 0 [m], 0 [m]")
                        outStr.NewLines 1
                    Case "Surface group"
                        If templateString = "" Then templateString = GetRange("Template.SurfaceGroup")  ' Load default if no template given
                        templateString = ReplaceMultiple(templateString, userArgs)
                        outStr.AppendRow ReplaceMultiple(templateString, "${NAME}", userLocations(i, 1), "${LOCATION_LIST}", firstUserLocation, "${TRANSPARENCY}", "0.2")
                        outStr.NewLines 1
                    Case "Contour"
                        If templateString = "" Then templateString = GetRange("Template.Contour")  ' Load default if no template given
                        templateString = ReplaceMultiple(templateString, userArgs)
                        outStr.AppendRow ReplaceMultiple(templateString, "${NAME}", userLocations(i, 1), "${VARIABLE}", "Pressure", "${LOCATION_LIST}", firstUserLocation)
                        outStr.NewLines 1
                    Case "Streamline"
                        If templateString = "" Then templateString = GetRange("Template.Streamline")  ' Load default if no template given
                        templateString = ReplaceMultiple(templateString, userArgs)
                        outStr.AppendRow ReplaceMultiple(templateString, "${NAME}", userLocations(i, 1))
                        outStr.NewLines 1
                End Select
            End If
        Next i
    Next objectType
    
    OutputUserLocationsAndPlots = outStr.Output
    
    Exit Function
ErrorHandler:
    If Err.Number = vbObjectError + 513 Then
        MsgBox Err.Description, vbCritical, "Range not found"
        Resume Next
    End If
    
End Function


Function OutputUserLocationsAndPlots() As String
' Action: Returns the cmdstr to create the user locations (Planes, surface groups, etc) and plots (Contour, Streamline, etc)
'
    Dim outStr As New ClassString
    
    Dim i As Integer, j As Integer
    Dim userLocations As Range
    Set userLocations = GetRange("UserLocations")
    Dim firstUserLocation As String
    firstUserLocation = "inlet"
    
    Dim userLocationDefaults As Range, objectType As String, defaultArgs As Variant
    Dim defaultTemplateName As String, defaultTemplateString As String
    Set userLocationDefaults = GetRange("UserLocationDefaults")
    
    On Error GoTo ErrorHandler
    
    ' Loop through all userLocation objects
    For i = 1 To userLocationDefaults.Rows.Count
        objectType = userLocationDefaults(i, 1)
        If objectType = "" Then GoTo NextObjectType
        
        defaultTemplateName = userLocationDefaults(i, 2)
        defaultTemplateString = GetRange(defaultTemplateName)
        defaultArgs = Split(userLocationDefaults(i, 3), ";")
        
        ' Loop through userLocation range
        Dim customTemplateName As String, templateString As String, customArgs As Variant
        For j = 1 To userLocations.Rows.Count
            If userLocations(j, 2) <> objectType Then GoTo NextUserLocation
            If userLocations(i, 3) <> "" Then
                templateString = GetRange(userLocations(j, 3))
            Else
                templateString = defaultTemplateString
            End If
            customArgs = Split(userLocations(j, 4), ";")
            templateString = ReplaceMultiple(templateString, "${NAME}", userLocations(j, 1))
            templateString = ReplaceMultiple(templateString, customArgs)
            templateString = ReplaceMultiple(templateString, defaultArgs)
            outStr.AppendRow templateString
            outStr.NewLines 1
NextUserLocation:
        Next j
NextObjectType:
    Next i
    
    OutputUserLocationsAndPlots = outStr.Output
    Exit Function
ErrorHandler:
    If Err.Number = vbObjectError + 513 Then
        MsgBox Err.Description, vbCritical, "Range not found"
        Resume Next
    ElseIf Err.Number = vbObjectError + 514 Then
        Err.Raise vbObjectError + 514, Err.Source, Err.Description
    End If
    
End Function


Function OutputUserLocation(userLocation As Range) As String

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
                userLocName = Trim(userLocs(j))
                Dim r As Range
                Set r = GetRow("UserLocations", 1, userLocName)
                If Not r Is Nothing Then
                    userLocType = r.Cells(1, 2)
                Else
                    Dim question
                    question = MsgBox("Error in figure """ & targetFigure & """, user location """ & userLocName & """ not found", vbExclamation + vbYesNoCancel, "User location not found")
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

