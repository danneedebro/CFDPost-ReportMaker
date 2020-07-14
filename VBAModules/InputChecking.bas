Attribute VB_Name = "InputChecking"
Option Explicit
Option Private Module

Private Enum enum_check_input_return_value
    input_error_continue_true
    input_error_continue_false
    OK
End Enum

Function CheckInput() As Boolean
' Action: Check input
'
    Dim input_check_ok As Boolean, answ As Variant
    input_check_ok = True
    
    ' Check if main input tables exists
    If CheckInputTables() = False Then
        CheckInput = False
        Exit Function
    End If
    
    
    ' Check figures
    Dim check_figures As enum_check_input_return_value
    check_figures = CheckInputFigures()
    If check_figures = input_error_continue_false Then
        CheckInput = False
        Exit Function
    ElseIf check_figures = input_error_continue_true Then
        input_check_ok = False
    End If
    
    ' Check views
    Dim check_views As enum_check_input_return_value
    check_views = CheckInputViews()
    If check_views = input_error_continue_false Then
        CheckInput = False
        Exit Function
    ElseIf check_views = input_error_continue_true Then
        input_check_ok = False
    End If
    
    CheckInput = input_check_ok
    
End Function

Private Function CheckInputTables() As Boolean
' Action: Check if main input tables exists
'
    Dim input_check_ok As Boolean
    input_check_ok = True
    
    Dim table As Variant, table_range As Range, tables_missing As String
    For Each table In Array("Figures.Geometry", "Figures.Mesh", "Figures.Solution", "Figures.External.Convergence", _
                            "Figures.External.Misc", "Views", "UserLocations")
        Set table_range = GetRange(CStr(table), False)
        If table_range Is Nothing Then
            tables_missing = tables_missing & table & vbNewLine
            input_check_ok = False
        End If
    Next table
    
    If input_check_ok = False Then
        MsgBox "Error: Can't find input tables: " & vbNewLine & vbNewLine & tables_missing, vbExclamation, "Can't find main input tables"
    End If
    
    CheckInputTables = input_check_ok
    
End Function

Private Function CheckInputFigures() As enum_check_input_return_value
' Action: Check figure input
'
    Dim input_check_ok As Boolean, answ As Variant
    input_check_ok = True

    ' Check figures
    Dim i As Integer, figures As Range, table As Variant
    For Each table In Array("Figures.Geometry", "Figures.Mesh", "Figures.Solution")
        Set figures = GetRange(CStr(table), False)
        If figures Is Nothing Then
            input_check_ok = False
            answ = MsgBox("Error: Can't find figure table named """ & CStr(table) & """. Continue checking input?", vbExclamation + vbYesNoCancel, "Can't find view custom template")
            If answ <> vbYes Then GoTo InputErrorQuit Else GoTo NextFigureTable
        End If
        
        For i = 1 To figures.Rows.Count
            Dim figure_name As String, figure_view_range As Range, figure_view_name As String, figure_user_loc_and_plots As String
            figure_name = figures(i, 1).Text
            figure_view_name = figures(i, 2).Text
            figure_user_loc_and_plots = figures(i, 4)
            
            ' Check if view exists
            Set figure_view_range = FindRowInRange("Views", 1, figure_view_name)
            If figure_view_range Is Nothing Then
                input_check_ok = False
                figures(i, 2).Select
                answ = MsgBox("Error: Can't find view named """ & figure_view_name & """. Continue checking input?", vbExclamation + vbYesNoCancel, "Can't find View")
                If answ <> vbYes Then GoTo InputErrorQuit
            End If
            
            ' Check user location or plot object exists
            Dim user_loc_or_plot_name As Variant, user_loc_or_plot_range As Range
            For Each user_loc_or_plot_name In Split(figure_user_loc_and_plots, ",")
                Set user_loc_or_plot_range = FindRowInRange("UserLocations", 1, CStr(Trim(user_loc_or_plot_name)))
                If user_loc_or_plot_range Is Nothing Then
                    input_check_ok = False
                    figures(i, 4).Select
                    answ = MsgBox("Error: Can't find user location or plot named """ & user_loc_or_plot_name & """. Continue checking input?", vbExclamation + vbYesNoCancel, "Can't find User location or plot")
                    If answ <> vbYes Then GoTo InputErrorQuit
                End If
            Next user_loc_or_plot_name
        Next i
NextFigureTable:
    Next table
    
' Return results
    If input_check_ok = True Then
        CheckInputFigures = OK
    Else
        CheckInputFigures = input_error_continue_true
    End If
    Exit Function
    
InputErrorQuit:
    CheckInputFigures = input_error_continue_false
    
    
End Function

Private Function CheckInputViews() As enum_check_input_return_value
' Action: Check input for views
'
    Dim input_check_ok As Boolean, answ As Variant
    input_check_ok = True

    ' Check views
    Dim views As Range, view_name As String, view_user_template As String, view_user_args As String, view_range As Range
    Dim i As Integer
    Set views = GetRange("Views")
    For i = 1 To views.Rows.Count
        view_name = views(i, 1)
        view_user_template = views(i, 3)
        view_user_args = views(i, 4)
        If view_user_template <> "" Then
            Set view_range = GetRange(view_user_template, False)
            If view_range Is Nothing Then
                input_check_ok = False
                views(i, 3).Select
                answ = MsgBox("Error: Can't find custom template named """ & view_user_template & """. Continue checking input?", vbExclamation + vbYesNoCancel, "Can't find view custom template")
                If answ <> vbYes Then GoTo InputErrorQuit
            End If
        End If
    Next i
        
    ' Return results
    If input_check_ok = True Then
        CheckInputViews = OK
    Else
        CheckInputViews = input_error_continue_true
    End If
    Exit Function
    
InputErrorQuit:
    CheckInputViews = input_error_continue_false
    
End Function

Private Function CheckInputUserLocationsAndPlots() As enum_check_input_return_value
' Action: Checks the table UserLocationsAndPlots
'
    Dim input_check_ok As Boolean, answ As Variant
    input_check_ok = True

    Dim user_locations_and_plots As Range
    Set user_locations_and_plots = GetRange("UserLocations")
    If user_locations_and_plots Is Nothing Then
        answ = MsgBox("Error: Can't find user-location-and-plots table named ""UserLocations"". Continue?", vbExclamation + vbYesNoCancel, "Can't find user-location-and-plots table")
        If answ <> vbYes Then
            GoTo InputErrorQuit
        Else
            CheckInputUserLocationsAndPlots = input_error_continue_true
            Exit Function
        End If
    End If
    
    ' Loop through all User locations and plots
    Dim i As Integer
    For i = 1 To user_locations_and_plots.Rows.Count
        
    Next i
    

    ' Return results
    If input_check_ok = True Then
        CheckInputUserLocationsAndPlots = OK
    Else
        CheckInputUserLocationsAndPlots = input_error_continue_true
    End If
    Exit Function
    
InputErrorQuit:
    CheckInputUserLocationsAndPlots = input_error_continue_false
    
End Function

