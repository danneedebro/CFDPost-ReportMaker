Attribute VB_Name = "CFDPost_cmdlineGen"
Option Explicit

Private Enum enum_figure_type
    Geometry
    Mesh
    Solution
End Enum

Private Sub Text2Clipboard(OutputString As String)
'
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    MsgBox "Text copied. Start CFDPost and paste into command editor." & vbNewLine & vbNewLine & OutputString, vbOKOnly + vbInformation, "Command line arguments copied"
    
    With MSForms_DataObject
        .SetText OutputString
        .PutInClipboard
    End With

End Sub



Private Function GetPathAbsolute(Path_relative_or_absolute As String) As String
' Action: Combines (if necessary), the base path and the supplied path
'         TODO: Add more checks here to know if absolute or relative path is given
'
    Dim path_base As String
    path_base = Range("Path.Base")
    GetPathAbsolute = path_base & Path_relative_or_absolute
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

Private Function GetFigures(figure_type As enum_figure_type, Optional OnlyVisible As Boolean = True) As String
' Action: Returns a list of the figures
'
    Dim figures_range As Range
    If figure_type = Geometry Then
        Set figures_range = Range("Figures.Geometry")
    ElseIf figure_type = Mesh Then
        Set figures_range = Range("Figures.Mesh")
    ElseIf figure_type = Solution Then
        Set figures_range = Range("Figures.Solution")
    End If
    
    Dim i As Integer, figures_list As String
    For i = 1 To figures_range.Rows.Count
        If figures_range(i, 3).Value = "Yes" Or OnlyVisible = False Then
            figures_list = figures_list & "/VIEW:" & figures_range(i, 1).Value & ","
        End If
    Next i
    figures_list = Left(figures_list, Len(figures_list) - 1) ' Remove last comma
    GetFigures = figures_list
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
    
    Dim i As Integer, outstr As String
    For i = LBound(args) To UBound(args)
        outstr = outstr + CStr(args(i)) & IIf(i < UBound(args), ";", "")
    Next i
    ArgList = outstr
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
    Set userLocationDefault = FindRowInRange("UserLocationDefaults", 1, objectType)
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
    Dim inp As Variant, i As Integer, outstr As String
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
        outstr = outstr & inps(i) & IIf(i = inps.Count, "", ",")
    Next i
    
    userLocation.Cells(1, 4).Formula = "=ArgList(" & outstr & ")"
        
End Sub

Private Sub TTEST()
    Debug.Print CheckRange("Views")
End Sub


Private Function CheckRange(RangeName As String, Optional Wrksht As Worksheet) As Boolean
' Action: Check if range is OK
'
    On Error GoTo ErrorHandler
    Dim range_to_test As Range
    Set range_to_test = Range(RangeName)
    CheckRange = True
    
    Exit Function
ErrorHandler:
    CheckRange = False
    
End Function




Sub Test_CheckInput()
    Debug.Print CheckInput()
End Sub

Sub CreateReportSkeleton()
' Action: Creates the skeleton for the state file (.cst)
'
    Dim outstr As New ClassString, i As Integer
    
    If CheckInput() = False Then Exit Sub
    
    ' STEP 0: Load result file
    outstr.AppendRow OutputSub_HideShowObjects()
    
    outstr.AppendRow "!sub LoadResultFile{"
    outstr.AppendRow ">close"
    outstr.AppendRow OutputLoadFile()
    outstr.AppendRow "!}"
    
    ' STEP 1: Create user locations and plots
    outstr.AppendRow "!sub CreateUserLocationsAndPlots{"
    outstr.AppendRow OutputUserLocationsAndPlots()
    outstr.AppendRow "!}"
    
    ' STEP 2: Create model description
    outstr.AppendRow "!sub UpdateModelDescription{"
    outstr.AppendRow OutputModelDescription()
    outstr.AppendRow "!}"
    
    ' STEP 3: Create Result table
    outstr.AppendRow "!sub UpdateResultTable{"
    outstr.AppendRow "!# Action: Updates result table" & vbNewLine & "!#"
    outstr.AppendRow OutputResultTable()
    outstr.AppendRow "!}" & vbNewLine
    
    outstr.AppendRow OutputSub_CreateViews()
    
    ' STEP 4: Create figures and report elements
    outstr.AppendRow OutputSub_CreateFigures()
        
    ' STEP 5: Reset views
    outstr.AppendRow "!sub ResetViews{"
    outstr.AppendRow OutputResetViews()
    outstr.AppendRow "!}"
    
    ' STEP 5: Create external plots
    outstr.AppendRow "!sub LoadExternalFigures{"
    outstr.AppendRow Indent(OutputExternalFigures())
    outstr.AppendRow "!}"
    
    ' STEP 6: Update report settings
    outstr.AppendRow OutputSub_ReportSettings()
    
    ' STEP 7: Set report order
    outstr.AppendRow OutputSub_ReportSortItems()
    
    ' STEP 8: Publish report
    outstr.AppendRow "!sub PublishReport{"
    outstr.AppendRow "    REPORT:"
    outstr.AppendRow "       PUBLISH:"
    outstr.AppendRow "          Report Path = $_[0]"
    outstr.AppendRow "       END"
    outstr.AppendRow "    END"
    outstr.AppendRow "    > update"
    outstr.AppendRow "    >report save"
    outstr.AppendRow "!}"
    
    outstr.NewLines 2
    outstr.AppendRow "# Comment out the subroutines not to run"
    outstr.AppendRow "# Step 1. Load result and create"
    outstr.AppendRow "!LoadResultFile();"
    outstr.AppendRow "!CreateUserLocationsAndPlots();"
    outstr.AppendRow "!CreateFigures();"
    outstr.AppendRow "!ReportSettings();"
    outstr.AppendRow "# Step 2. Manually adjust User locations and plots"
    outstr.AppendRow "# Step 3. Manually adjust the camera of View 1 to View 4"
    outstr.AppendRow "# Step 4. Comment out subs in Step 1 and run subs below"
    outstr.AppendRow "!ResetViews();"
    outstr.AppendRow "!UpdateModelDescription();"
    outstr.AppendRow "!UpdateResultTable();"
    outstr.AppendRow "!LoadExternalFigures();"
    outstr.AppendRow "!ReportSortItems();"
    
    outstr.AppendRow "# Step 5: Publish report"
    outstr.AppendRow ReplaceMultiple("# !PublishReport(""${PATH}"");", "${PATH}", GetPathAbsolute(Range("Report.Path")))

    Text2Clipboard outstr.output
End Sub

Function OutputModelDescription() As String
' Action: Returns text for the model description
'
    Dim tmpStr As String, outstr As String
    
    tmpStr = "<p><b>Solver:</b><br>" & Range("Solver.Type").Value & ", " & Range("Solver.Time").Value & "</p>"
    tmpStr = tmpStr & "<p><b>Turbulence: " & "</b><br>Model = " & Range("TurbulenceModel.Name") & "<br>Wall function = " & Range("TurbulenceModel.WallFunction").Text & "</p>"
    tmpStr = tmpStr & "<p><b>Fluid: " & Range("Fluid.Description") & "</b><br>Density = " & Range("Fluid.Density") & " kg/m3<br>Viscosity = " & Range("Fluid.Viscosity").Text & " Pa·s</p>"
    tmpStr = tmpStr & ReplaceMultiple(Range("Template.CommentSubheading").Text, "${TITLE}", "Inlet:", "${TEXT}", Range("BC.Inlet").Text)
    tmpStr = tmpStr & ReplaceMultiple(Range("Template.CommentSubheading").Text, "${TITLE}", "Outlet:", "${TEXT}", Range("BC.Outlet").Text)
    tmpStr = tmpStr & ReplaceMultiple(Range("Template.CommentSubheading").Text, "${TITLE}", "Notes:", "${TEXT}", Replace(Range("Misc.Notes").Text, Chr(10), "<BR>")) & vbNewLine
    OutputModelDescription = ReplaceMultiple(Range("Template.Comment").Text, "${NAME}", "Header Description", "${COMMENT_HEADING_LEVEL}", "1", "${COMMENT_HEADING}", "Model description", "${COMMENT_TEXT}", tmpStr)
End Function

Function OutputResultTable() As String
' Action: Returns the result table
'
    Dim outstr As New ClassString, i As Integer, j As Integer, tableInput As Range
    Dim cellName As String
    outstr.AppendRow "   TABLE:Result Table"
    outstr.AppendRow "      TABLE CELLS:"
    Set tableInput = Range("TableInput")
    For i = 1 To tableInput.Rows.Count
        For j = 1 To tableInput.Columns.Count
            cellName = Chr(64 + j) & CStr(i)
            'If tableInput(i, j) <> "" Then
                outstr.AppendRow ReplaceMultiple("         ${CELL} = ""${FORMULA}"", False, False, False, Left, True, 0, Font Name, 1|1, %10.3e, True, ffffff, 000000, True", _
                                                 "${CELL}", cellName, "${FORMULA}", tableInput(i, j))
            'End If
        Next j
    Next i
        
    outstr.AppendRow "      END"
    outstr.Append "   END"
    
    OutputResultTable = outstr.output

End Function

Function OutputReportOrder() As String
' Action: Returns cmdstr for setting the report order
'
    Dim outstr As New ClassString
    outstr.Append "REPORT:" & vbNewLine & "  Report Items = /TITLE PAGE,/REPORT/FILE INFORMATION OPTIONS,/REPORT/MESH STATISTICS OPTIONS," & _
                  "/REPORT/PHYSICS SUMMARY OPTIONS,/REPORT/SOLUTION SUMMARY OPTIONS,/REPORT/OPERATING MAPS," & _
                  "/COMMENT:Header Description,/COMMENT:User Data,/TABLE:Result Table"
    
    ' Loop through figure tables (Geometry, Mesh and Solution)
    Dim figureTables As Variant, figureTable As Variant, figureTableRng As Range, i As Integer
    figureTables = Array("Geometry", "Mesh", "Solution")
    
    For Each figureTable In figureTables
        Set figureTableRng = Range("Figures." & figureTable)
        outstr.Append ",/COMMENT:Header " & figureTable
        
        For i = 1 To figureTableRng.Rows.Count
            If figureTableRng(i, 1).Text <> "" Then
                outstr.Append ",/VIEW:" & figureTableRng(i, 1)
            End If
        Next i
    Next figureTable
    
    outstr.Append ",/COMMENT:Header Convergence,/COMMENT:Header Misc" & vbNewLine & "END"
    
    OutputReportOrder = outstr.output
    
End Function

Function OutputExternalFigures() As String
' Action: Returns text for external figures
'
    ' External figures - convergence plots
    Dim figureTables As Variant, figureTable As Variant, figureTableRng As Range, i As Integer
    Dim tmpStr As String, outstr As String
    figureTables = Array("Convergence", "Misc")
    
    For Each figureTable In figureTables
        Set figureTableRng = Range("ExternalFigures." & figureTable)
        
        tmpStr = ""
        For i = 1 To figureTableRng.Rows.Count
            If figureTableRng(i, 3) <> "" Then
                Dim figure_path As String
                figure_path = GetPathAbsolute(figureTableRng(i, 3).Text)
            
                tmpStr = tmpStr & ReplaceMultiple(Range("Template.ExternalFigure").Text, "${PATH}", figure_path, Chr(10), "\" & Chr(10), _
                                                  "${FIGURE_NAME}", "Figure C" & CStr(i), "${FIGURE_CAPTION}", figureTableRng(i, 2).Text)
            End If
        Next i
        outstr = outstr & ReplaceMultiple(Range("Template.Comment").Text, "${NAME}", "Header " & figureTable, "${COMMENT_HEADING_LEVEL}", "1", "${COMMENT_HEADING}", figureTable, "${COMMENT_TEXT}", tmpStr)
        outstr = outstr & "" & vbNewLine
    Next figureTable
    
    OutputExternalFigures = outstr
End Function

Private Function OutputSub_HideShowObjects() As String
' Action: Outputs the Pearl subroutine to hide and show objects (user location an plots)
'
    Dim userLocations As Range
    Set userLocations = GetRange("UserLocations")
    Dim userlocation_list As String
    userlocation_list = ""
    
    Dim i As Integer, j As Integer
    For i = 1 To userLocations.Rows.Count
        Dim userlocation_type As String, userlocation_name As String
        userlocation_name = userLocations(i, 1)
        userlocation_type = userLocations(i, 2)
        If userlocation_name <> "" Then
            userlocation_list = userlocation_list & """/" & UCase(userlocation_type) & ":" & userlocation_name & ""","
        End If
    Next
    userlocation_list = Left(userlocation_list, Len(userlocation_list) - 1)
    
    OutputSub_HideShowObjects = ReplaceMultiple(Range("Template.HelperSubs"), "${USER_LOCATIONS}", userlocation_list)

End Function

Private Function OutputSub_ReportSettings() As String
' Action: Creates Sub that changes report settings
'
    Dim outstr As New ClassString
    Dim report_path As String
    report_path = GetPathAbsolute(Range("Report.Path"))
    
    outstr.AppendRow "!sub ReportSettings{"
    outstr.AppendRow "# Action: Sets report settings"
    Dim CCL_string As String
    CCL_string = ReplaceMultiple(Range("Template.ReportSettings").Text, "${FIGURE_HEIGHT}", Range("Figure.Height"), "${FIGURE_WIDTH}", _
                                      Range("Figure.Width"), "${REPORT_PATH}", report_path, "${REPORT_AUTHOR}", Range("Report.Author").Text, _
                                      "${REPORT_TITLE}", Range("Report.Title").Text)
    outstr.AppendRow Indent(CCL_string)
    outstr.AppendRow "!}"
    OutputSub_ReportSettings = outstr.output
End Function

Private Function OutputSub_ReportSortItems() As String
' Action: Creates Sub that sorts the report
'
    Dim outstr As New ClassString
    
    outstr.AppendRow "!sub ReportSortItems{"
    outstr.AppendRow "# Action: Sorts the figures in the report"
    outstr.AppendRow "# Template.Report.SortItems"
    Dim CCL_string As String
    CCL_string = ReplaceMultiple(Range("Template.Report.SortItems"), "${FIGURES_GEOMETRY}", GetFigures(Geometry), _
                              "${FIGURES_MESH}", GetFigures(Mesh), "${FIGURES_SOLUTION}", GetFigures(Solution))
    outstr.AppendRow Indent(CCL_string)
    outstr.AppendRow "!}"
    OutputSub_ReportSortItems = outstr.output
End Function

Private Function OutputSub_CreateViews() As String
' Action: Sub that creates the views
'
    Dim outstr As New ClassString
    outstr.AppendRow "!sub CreateViews(){"
    outstr.AppendRow "# Action: Create views"
    
    Dim default As Range, default_template_name As String, default_args As Variant
    Set default = GetRange("Defaults.Views")
    default_template_name = default.Cells(1, 2).Text
    default_args = Split(default.Cells(1, 3), ";")
    
    Dim i As Integer, views As Range, CCL_string As String, template_name As String, template_default As Boolean
    Dim view_name As String, view_user_template As String, view_user_args As Variant
    
    Set views = GetRange("Views")
    
    ' Loop through each view. If no user template given - use default.
    For i = 1 To views.Rows.Count
        CCL_string = ""
        view_name = views(i, 1)
        view_user_template = views(i, 3)
        view_user_args = Split(views(i, 4), ";")
        
        ' If no user template given
        If view_user_template = "" Then
            template_name = default_template_name
            CCL_string = ReplaceMultiple(GetRange(default_template_name), view_user_args)
            CCL_string = ReplaceMultiple(CCL_string, default_args)
        ' User template given
        Else
            template_name = views(i, 3)
            CCL_string = ReplaceMultiple(GetRange(template_name), view_user_args)
        End If
        
        CCL_string = ReplaceMultiple(CCL_string, "${NAME}", view_name)
        outstr.AppendRow "#   View: " & view_name & " (Template = " & template_name & ")"
        outstr.AppendRow Indent(CCL_string) & vbNewLine
    Next i
    outstr.AppendRow "!}"
    OutputSub_CreateViews = outstr.output
End Function


Private Function OutputSub_CreateFigures() As String
' Action: Creates the Perl sub that first creates all Figures
'
    On Error GoTo ErrorHandler
    Dim outstr As New ClassString
    
    Dim figureTables As Variant, figureTable As Variant, figureTableRng As Range, copy_type As String
    figureTables = Array("Geometry", "Mesh", "Solution")
    outstr.AppendRow "!sub CreateFigures{"
    outstr.AppendRow "!# Action: Creates figures"
    outstr.AppendRow "!#"
    
    If GetRange("Options.ShallowCopy") = "YES" Then
        copy_type = "shallow_copy"
    Else
        copy_type = "copy"
        outstr.AppendRow "!   # Activate all known objects to make a deep copy   "
        outstr.AppendRow "    > setViewportView cmd=set, view=/VIEW:View 1, viewport=1"
        outstr.AppendRow "!   HideShowObjects(""View 1"",""show"");"
    End If
    
    
    For Each figureTable In figureTables
        outstr.AppendRow Indent(ReplaceMultiple(Range("Template.Comment"), "${NAME}", "Header " & figureTable, "${COMMENT_HEADING_LEVEL}", "1", _
                                         "${COMMENT_HEADING}", figureTable, "${COMMENT_TEXT}", ""))
        Set figureTableRng = Range("Figures." & figureTable)
        
        Dim i As Integer
        For i = 1 To figureTableRng.Rows.Count
            If figureTableRng(i, 1).Text <> "" Then
                outstr.AppendRow ReplaceMultiple("    >delete /VIEW:${FIGURE}", "${FIGURE}", figureTableRng(i, 1))
                outstr.AppendRow ReplaceMultiple("    > setViewportView cmd=${COPY_TYPE}, view=/VIEW:${FIGURE}, viewport=1", "${FIGURE}", figureTableRng(i, 1), "${COPY_TYPE}", copy_type)
                ' TODO: HIDE LOCAL OBJECTS WITH This outStr.AppendRow "!   HideShowObjects(""View 1"",""hide"");"
            End If
        Next i
    Next figureTable
    outstr.AppendRow "!}"
    
    OutputSub_CreateFigures = outstr.output
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    
End Function



Function OutputUserLocationsAndPlots() As String
' Action: Returns the cmdstr to create the user locations (Planes, surface groups, etc) and plots (Contour, Streamline, etc)
'
    Dim outstr As New ClassString
    
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
            outstr.AppendRow templateString
            outstr.NewLines 1
NextUserLocation:
        Next j
NextObjectType:
    Next i
    
    OutputUserLocationsAndPlots = outstr.output
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
    Dim outstr As New ClassString
    
    Dim i As Integer, j As Integer
    'Dim outStr As String
    'outStr = ""
    
    ' Define HideAllKnownObjects subroutine
    Dim userLocations As Range
    Set userLocations = Range("UserLocations")
    Dim locType As String
    outstr.AppendRow "# Hides all known objects in Figure X"
    outstr.AppendRow "!   sub HideAllKnownObjects {"
    For i = 1 To userLocations.Rows.Count
        If userLocations(i, 1) <> "" Then
            outstr.AppendRow "      >hide /" & UCase(userLocations(i, 2)) & ":" & userLocations(i, 1) & ", view=/VIEW:$_[0]"
        End If
    Next i
    outstr.AppendRow "!   }" & vbNewLine
    
    Dim figures As Range, figureTables As Variant, figureTable As Variant
    Dim sourceView As String, targetFigure As String, userLocType As String, userLocName As String
    Dim userLocs As Variant
    figureTables = Array("Figures.Geometry", "Figures.Mesh", "Figures.Solution")
    
    ' Loop through each figureTable (Geometry, Mesh and Solution) and hide all objects then
    ' show the ones that suppose to be shown
    For Each figureTable In figureTables
        Set figures = GetRange(CStr(figureTable))
        For i = 1 To figures.Rows.Count
            sourceView = figures(i, 2)
            targetFigure = figures(i, 1)
            outstr.AppendRow "!   # Figure: " & targetFigure
            outstr.AppendRow "!   HideAllKnownObjects(""" & targetFigure & """);"
            userLocs = Split(figures(i, 4), ",")
            
            For j = LBound(userLocs) To UBound(userLocs)
                userLocName = Trim(userLocs(j))
                Dim r As Range
                Set r = FindRowInRange("UserLocations", 1, userLocName)
                If Not r Is Nothing Then
                    userLocType = r.Cells(1, 2)
                Else
                    Dim question
                    question = MsgBox("Error in figure """ & targetFigure & """, user location """ & userLocName & """ not found", vbExclamation + vbYesNoCancel, "User location not found")
                    If question <> vbYes Then Exit Function Else userLocType = "Plane"
                End If
                outstr.AppendRow "    >show /" & UCase(userLocType) & ":" & userLocName & ", view=/VIEW:" & targetFigure
            Next j
            
            
            outstr.AppendRow "    >delete /VIEW:" & targetFigure & "/CAMERA"
            outstr.AppendRow "    >copy from = /VIEW:" & sourceView & "/CAMERA, to = /VIEW:" & targetFigure & "/CAMERA"
            If figures(i, 3) = "Yes" Then
                outstr.AppendRow "    > report showItem=/VIEW:" & targetFigure
            Else
                outstr.AppendRow "    > report hideItem=/VIEW:" & targetFigure
            End If
            
            outstr.AppendRow ""
            
        Next i
    Next figureTable
    
    OutputResetViews = outstr.output
    
End Function

Function OutputReportSettings() As String
' Action: Returns cmdstr for setting the report settings
'
    Dim outstr As New ClassString
    Dim report_path As String
    report_path = GetPathAbsolute(Range("Report.Path"))

    ' Update report settings
    outstr.AppendRow ReplaceMultiple(Range("Template.ReportSettings").Text, "${FIGURE_HEIGHT}", Range("Figure.Height"), "${FIGURE_WIDTH}", _
                                      Range("Figure.Width"), "${REPORT_PATH}", report_path, "${REPORT_AUTHOR}", Range("Report.Author").Text, _
                                      "${REPORT_TITLE}", Range("Report.Title").Text)
    ' White background
    outstr.AppendRow ">setPreferences Viewer Background Colour Type = Solid, Viewer Background Image File =  , Viewer Background Colour = 1.0&1.0&1.0"
    
    OutputReportSettings = outstr.output
    
End Function

Function OutputLoadFile() As String
' Action: Returns the cmdstr for loading a file
'
    Dim outstr As New ClassString
    outstr.AppendRow ReplaceMultiple(Range("Template.LoadFile"), "${FILENAME}", GetPathAbsolute(Range("ResultFile")))
    
    OutputLoadFile = outstr.output
    
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



' -----------------------------------------------TESTING----------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------------------------
Private Sub Test_GetFigures()
    Debug.Print GetFigures(Geometry)
    Debug.Print GetFigures(Mesh)
    Debug.Print GetFigures(Solution)
    Debug.Print GetFigures(Solution, False)
End Sub

Private Sub Test_CreateViews()
    Debug.Print OutputSub_CreateViews()
End Sub

Private Sub Test_CreateFigures()
    Debug.Print OutputSub_CreateFigures()
End Sub
