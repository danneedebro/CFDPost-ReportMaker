VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlotObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public PlotObjectType As String
Public Template As TemplateObject
Private pLog As String

Public Property Get Log(Optional IndentLevel As Integer = 0) As String
    Dim indent_str As String
    indent_str = String(IndentLevel * 3, " ")
    Log = Replace(indent_str & pLog, Find:=vbNewLine, Replace:=vbNewLine & indent_str)
End Property

Public Property Let Log(Optional IndentLevel As Integer = 0, arg As String)
    pLog = arg
End Property

Private Sub Class_Initialize()
    Set Template = New TemplateObject
End Sub

Private Sub ResetMe()
    pLog = TypeName(Me) & " object created" & vbNewLine
End Sub

Public Function Create(Name As String, PlotObjectType As String, TemplateRangeAddress As String, Arguments As String) As Boolean
' Action: Creates a range
'
    On Error GoTo ErrorHandler
    Dim success As Boolean, section_id As String
    success = True
    
    section_id = "Reset object"
    ResetMe

    ' Set basic properties
    section_id = "Basic properties"
    AppendLog "PlotObject.Create(Name:=""" & Name & """,PlotObjectType:=""" & PlotObjectType & """,TemplateRangeAddress:=""" & TemplateRangeAddress & """,Arguments:=""" & Arguments & """)"
    Me.Name = Name
    Me.PlotObjectType = PlotObjectType
    Arguments = "${NAME};" & Name & IIf(Arguments = "", "", ";" & Arguments)
    
    ' Update template object properties
    section_id = "Read template"
    Me.Template.Arguments = Arguments
    Me.Template.RangeAddress = TemplateRangeAddress
    
    If Me.Refresh() = False Then
        Create = False
        Exit Function
    Else
        Create = True
    End If
    
    
    Exit Function
ErrorHandler:
    Create = False
    AppendLog "Unknown error in " & TypeName(Me) & ".Create() in section = """ & section_id & """"
End Function

Public Function Refresh(Optional ReloadTemplateFromRange As Boolean = True) As Boolean
' Action: Refreshes plot with current properties
'
' Arguments:
'    ReloadTemplateFromRange = if true the property .TextWithWildcards is reloaded from range
'
' Output:
'    True if success, False otherwise
'
    On Error GoTo ErrorHandler
    Dim success_flag As Boolean, section_id As String
    success_flag = True
    
    If Me.Template.Refresh(ReloadTemplateFromRange) = False Then
        Refresh = False
        AppendLog "Failed reading template object:" & vbNewLine & Me.Template.Log(1)
        Exit Function
    End If
    
    Refresh = True
    
    Exit Function

' Catch unknown errors
ErrorHandler:
    Refresh = False
    AppendLog "Unknown error in " & TypeName(Me) & ".Refresh() in section = """ & section_id & """"
End Function

Public Function Copy() As PlotObject
' Action: Makes a copy of itself
'
    Dim output_object As New PlotObject
    With output_object
        .Name = Me.Name
        .PlotObjectType = Me.PlotObjectType
        Set .Template = Me.Template.Copy
    End With
        
    Set Copy = output_object
        
End Function

Private Sub AppendLog(NewText As String, Optional NumberOfNewLines As Integer = 1)
' Action: Appends the object log with NewText
'
    pLog = pLog & NewText & String(NumberOfNewLines, vbNewLine)
End Sub


