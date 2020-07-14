VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TABLE_FIGURES_GEOMETRY = "Figures.Geometry"
Private Const TABLE_FIGURES_MESH = "Figures.Mesh"
Private Const TABLE_FIGURES_SOLUTION = "Figures.Solution"
Private Const TABLE_PLOT_OBJECTS = "UserLocations"
Private Const TABLE_PLOT_OBJECT_DEFAULTS = "UserLocationDefaults"
Private Const TABLE_VIEWS = "Views"
Private Const TABLE_VIEW_DEFAULTS = "Views.Default"


Public PlotObjects As PlotObjectCollection




Private Sub Class_Initialize()
    Set Me.PlotObjects = New PlotObjectCollection

    SetPlotObjects
End Sub


Private Function GetPlotObjectDefaults() As PlotObjectCollection
' Action: Return a collection of plot object defaults
'
    ' Fetch table from spreadsheet
    Dim table As New TableObject
    If table.Create(TABLE_PLOT_OBJECT_DEFAULTS) = False Then
        MsgBox "Error: Failed reading table """ & TABLE_PLOT_OBJECT_DEFAULTS & """."
        Exit Function
    End If
    
    ' Loop through table and create default plot objects
    Dim i As Integer
    Dim plot_object As PlotObject, plot_objects As New PlotObjectCollection
    For i = 1 To table.NumberOfRows
        If table(i, 1) = "" Then GoTo NextRow
        Set plot_object = New PlotObject
        If plot_object.Create(Name:=table(i, 1), PlotObjectType:=table(i, 1), TemplateRangeAddress:=table(i, 2), Arguments:=table(i, 3)) = False Then
            table.SelectRow i
            MsgBox "Error: Failed reading plot object default, type =""" & table(i, 1) & """" & vbNewLine & plot_object.Log, vbExclamation, "Failed reading plot object default"
        Else
            plot_objects.Add plot_object
        End If
NextRow:
    Next i
    
    Set GetPlotObjectDefaults = plot_objects
    
    
End Function

Private Sub SetPlotObjects()
    ' Fetch table from spreadsheet
    Dim table As New TableObject
    If table.Create(TABLE_PLOT_OBJECTS) = False Then
        MsgBox "Error: Failed reading table """ & TABLE_PLOT_OBJECTS & """."
        Exit Sub
    End If
    
    ' Get plot object defaults
    Dim plot_object_defaults As PlotObjectCollection
    Set plot_object_defaults = GetPlotObjectDefaults()
    If plot_object_defaults Is Nothing Then
        MsgBox "Error: failed loading plot object defaults (table """ & TABLE_PLOT_OBJECT_DEFAULTS & """)"
        Exit Sub
    End If
    
    ' Loop through table and create a collection of plot objects. If no template is given, use the default one
    Dim i As Integer
    Dim plot_objects As New PlotObjectCollection
    Dim plot_object As PlotObject, plot_object_name As String, plot_object_type As String, plot_object_template As String, plot_object_args As String
    Dim plot_object_default As PlotObject
    
    For i = 1 To table.NumberOfRows
        plot_object_name = table(i, 1)
        plot_object_type = table(i, 2)
        plot_object_template = table(i, 3)
        plot_object_args = table(i, 4)
        plot_object_args = "${NAME};" & plot_object_name & IIf(plot_object_args = "", "", ";" & plot_object_args)
        
        If plot_object_name = "" Then GoTo NextRow
        
        ' If no template is given, use default template
        If plot_object_template = "" Then
            Set plot_object_default = plot_object_defaults.FindItemByName(plot_object_type)
            If plot_object_default Is Nothing Then
                MsgBox "Error: No default template for plot object type = """ & plot_object_type & """ found. Using empty"
                Set plot_object = New PlotObject
            Else
                Set plot_object = plot_object_default.Copy
                plot_object.Name = plot_object_name
                plot_object.Template.Arguments = plot_object_args & IIf(plot_object.Template.Arguments = "", "", ";" & plot_object.Template.Arguments)
                If plot_object.Refresh = False Then
                    MsgBox "Error: with"
                End If
            End If
        ' If template is given
        Else
            Set plot_object = New PlotObject
            If plot_object.Create(plot_object_name, plot_object_type, plot_object_template, plot_object_args) = False Then
                MsgBox "Error: Failed to create plot object"
                GoTo NextRow
            End If
        End If
        
        plot_objects.Add plot_object, plot_object_name
        
NextRow:
    Next i
    
    Set Me.PlotObjects = plot_objects
    
End Sub
