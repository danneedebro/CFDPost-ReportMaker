Attribute VB_Name = "Testing"
Option Explicit

Private Sub TestReplaceMultiple()

    Debug.Print ReplaceMultiple("${1} ${2}", "${1}", "Text1", "${2}", "Text2")
    
    Debug.Print ReplaceMultiple("${1} ${2}", Array("${1}", "Text1", "${2}", "Text2"))
    
    Debug.Print ReplaceMultiple("${1} ${2}", Split("${1};Text1;${2};Text2", ";"))
    
    
    Dim apa As Variant
    apa = Array()
    Debug.Print ReplaceMultiple("${1} ${2}", apa)
    
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





Private Sub Test_TemplateObject()
    Dim template_obj As New TemplateObject
    
    ' Assert that creation will fail if template doesn't exist
    Debug.Assert template_obj.Create("Template.DoesNotExist", "") = False
    
    ' Assert that creation will fail if arguments not a multiple of two
    Debug.Assert template_obj.Create("Template.Plane", "${Wildcard1};Value1;${Wildcard2}") = False
    
    ' This should display
    If template_obj.Create("Template.Plane", "${NAME};Plane 1") = True Then
        Debug.Print template_obj.TextWithWildcards
        Debug.Print template_obj.Text
        Debug.Print template_obj.Log(0)
    End If
    
    ' Update previous object
    template_obj.Arguments = "${NAME};Plane 1 (New)"
    If template_obj.Refresh(False) = True Then
        Debug.Print template_obj.Text
    End If
    
    ' Copy object
    Dim template_obj_new As TemplateObject
    Set template_obj_new = template_obj.Copy()
    Debug.Print template_obj_new.Log
    
End Sub


Private Sub Test_PlotObject()
    Dim plot_object As New PlotObject

    ' This should fail because template doesn't exist
    Debug.Assert (plot_object.Create("My plane", "Plane", "Template.DoesNotExist", "") = False)
    Debug.Print plot_object.Log

    ' This should succeed after updating template address
    plot_object.Template.RangeAddress = "Template.Plane"
    Debug.Assert (plot_object.Refresh = True)

    If plot_object.Create("My Plane", "Plane", "Template.Plane", "") = False Then
        Debug.Print "Fail"
        Debug.Print plot_object.Log
    Else
        Debug.Print plot_object.Template.Text
    End If
    
    ' Verify that it indeed creates a deep copy of the plot object
    Dim plot_object_copy As PlotObject
    Set plot_object_copy = plot_object.Copy
    plot_object_copy.Template.Arguments = ""
    Debug.Assert (plot_object_copy.Template.Arguments <> plot_object.Template.Arguments)
    
End Sub


Private Sub Test_PlotObjects()
    Dim plot_objects As New PlotObjectCollection

    Dim plot_object_1 As New PlotObject
    plot_object_1.Create "My Plane 1", "Plane", "Template.Plane", ""

    Dim plot_object_2 As New PlotObject
    plot_object_2.Create "My Plane 2", "Plane", "Template.Plane", ""

    plot_objects.Add plot_object_1, "Plane 1"
    plot_objects.Add plot_object_2, "Plane 2"

    Dim plot_object As PlotObject
    For Each plot_object In plot_objects
        Debug.Print plot_object.Name
    Next plot_object

    Dim obj_found As PlotObject
    Set obj_found = plot_objects.FindItemByName("My Plane 12")
    If obj_found Is Nothing Then
        Debug.Print "Object not found"
    Else
        Debug.Print "Object found"
    End If
    'Debug.Assert ((obj_found Is Nothing) = True)

    Debug.Print plot_objects("Plane 1").Name

End Sub


Private Sub Test_TableObject()

'    Dim apa As Range
'    Set apa = Range("Figures.Geometry")
'
'    Dim rowRng As Range
'        For Each rowRng In apa.Rows
'            Debug.Print rowRng(1, 4).Value
'        Next rowRng

    Dim table As New TableObject
    
    If table.Create("Figures.Geometry") = True Then
        Debug.Print table.Log

        'table.Item().1
        Dim rowRng As Range
        For Each rowRng In table.GetRange().Rows
            'Debug.Print rowRng.Cells(1, 1).Value
            Debug.Print rowRng.Value2(1, 1)
        Next rowRng

    Else

    End If
End Sub


Public Sub Test_GetPlotObjectDefaults()
    Dim table As New TableObject
    If table.Create("UserLocationDefaults") = False Then
        MsgBox "Error: Failed reading table ""UserLocationDefaults""."
        Exit Sub
    End If
    
    Dim i As Integer
    Dim plot_object As PlotObject, plot_objects As New PlotObjectCollection
    For i = 1 To table.NumberOfRows
        If table(i, 1) = "" Then GoTo NextRow
        Set plot_object = New PlotObject
        If plot_object.Create(Name:=table(i, 1), PlotObjectType:=table(i, 1), TemplateRangeAddress:=table(i, 2), Arguments:=table(i, 3)) = False Then
            table(i, 1).Select
            MsgBox "Failed reading """ & table(i, 1) & """" & vbNewLine & plot_object.Log
        Else
            plot_objects.Add plot_object
        End If
NextRow:
    Next i
    
    
    For Each plot_object In plot_objects
        Debug.Print plot_object.Log
    Next plot_object
End Sub


Sub Test_Main()

    Dim main_obj As Main
    Set main_obj = New Main
    
    Dim po As PlotObject
    For Each po In main_obj.PlotObjects
        Debug.Print po.Name
    Next po
    

End Sub



'Public Sub Test_s()
'    Dim i As Integer
'
'    ' Read templates and other input for plot object defaults
'    Dim table_defaults As New TableObject
'    If table_defaults.Create("UserLocationDefaults") = False Then
'    MsgBox "Error: Failed reading ""UserLocationDefaults""", vbCritical, "Error"
'        Exit Sub
'    End If
'
'    Dim plot_object_defaults As New PlotObjectCollection
'    Dim p As PlotObject
'
'    With table_defaults
'        For i = 1 To .NumberOfRows
'            Set p = New PlotObject
'            If p.Create(.Item(i, 1), .Item(i, 1), .Item(i, 2), .Item(i, 3)) = True Then
'                plot_object_defaults.Add p, p.PlotObjectType
'            End If
'        Next i
'    End With
'
'    ' Read templates and other input for plot objects
'    Dim table As New TableObject
'    If table.Create("UserLocations") = False Then
'        MsgBox "Error: Failed reading ""UserLocations""", vbCritical, "Error"
'        Exit Sub
'    End If
'
'
'    Dim plot_objects As New PlotObjectCollection
'
'    Dim plot_object_type_template As String, plot_object_type As String, plot_object_name As String, plot_object_args As String
'    Dim p_def As PlotObject
'    With table
'        For i = 1 To .NumberOfRows
'            plot_object_name = .Item(i, 1)
'            plot_object_type = .Item(i, 2)
'            plot_object_type_template = .Item(i, 3)
'
'            ' If no custom template is given, use default
'            If plot_object_type_template = "" Then
'                Set p_def = plot_object_defaults.FindItemByName(plot_object_type)
'                If p_def Is Nothing Then
'                    MsgBox "Default object not found in ""UserLocationDefaults"".", vbCritical, "Error"
'                Else
'                    Set p = New PlotObject
'                    If p.Create(plot_object_name, plot_object_type, p_def.RangeAddress, "hej") = False Then
'
'                    End If
'                    p.Name = plot_object_name
'                End If
'
'            ' If custom template is given
'            Else
'                Set p = New PlotObject
'                If p.Create(.Item(i, 1), .Item(i, 2), .Item(i, 3), .Item(i, 4)) = True Then
'                    plot_objects.Add p, p.Name
'                Else
'                    MsgBox "Error: Failed reading """ & p.Name & """" & vbNewLine & p.Log, vbCritical, "Error"
'                End If
'            End If
'
'
'        Next i
'    End With
'
'
'
'
'
'End Sub



'Private Sub TTEST()
'
'    Dim apa As New Collection
'
'    apa.Add "hej"
'    apa.Add "då"
'
'    Debug.Print Join(apa, ";")
'End Sub
