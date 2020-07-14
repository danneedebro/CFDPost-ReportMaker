VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplateObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Text As String
Public TextWithWildcards As String
Public Arguments As String
Public RangeAddress As String
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
    ResetMe
End Sub

Private Sub ResetMe()
    Me.Text = ""
    Me.TextWithWildcards = ""
    pLog = TypeName(Me) & " object created" & vbNewLine
End Sub

Public Function Create(TemplateRangeAddress As String, Arguments As String) As Boolean
' Action: Creates a template object
'
' Arguments:
'    TemplateRangeAddress = The address of where the template is stored
'    Arguments = A semicolon-separated string list with the arguments
'
' Output:
'    True if success, False otherwise
'
    On Error GoTo ErrorHandler
    Dim section_id As String
    
    section_id = "Reset object"
    ResetMe

    AppendLog "TemplateObject.Create(TemplateRangeAddress:=""" & TemplateRangeAddress & """,Arguments:=""" & Arguments & """)"
    Me.RangeAddress = TemplateRangeAddress
    Me.Arguments = Arguments
    
    ' Fetch template text
    section_id = "Fetch template"
    If Refresh() = False Then
        Create = False
        Exit Function
    Else
        Create = True
    End If
    
    Exit Function
ErrorHandler:
    Create = False
    AppendLog "Unknown error in " & TypeName(Me) & ".Refresh() in section = """ & section_id & """"
End Function


Public Function Refresh(Optional ReloadTemplateFromRange As Boolean = True) As Boolean
' Action: Refreshes template with current properties
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
    
    ' Try setting the range
    section_id = "Get range"
    If ReloadTemplateFromRange = True Then
        Dim template_range As Range
        ThisWorkbook.Activate
        Set template_range = Range(Me.RangeAddress)
        
        ' Set TextWithWildcards property
        Me.TextWithWildcards = template_range.Text
    End If
    
    ' Check and split arguments
    section_id = "Split arguments"
    Dim args As Variant
    args = Split(Me.Arguments, ";")
    If (UBound(args) - LBound(args) + 1) Mod 2 <> 0 Then
        Refresh = False
        AppendLog "Error: arguments not a multiple of 2"
        Exit Function
    End If
    
    ' Replace text
    section_id = "Replace text"
    Dim i As Integer, outstr As String
    outstr = Me.TextWithWildcards
    For i = LBound(args) To UBound(args) - 1 Step 2
        outstr = Replace(outstr, args(i), args(i + 1))
    Next i
    Me.Text = outstr
    
    Refresh = success_flag
    
    Exit Function

' Catch unknown errors
ErrorHandler:
    Refresh = False
    Select Case section_id
        Case "Get range"
            AppendLog "Failed setting range """ & Me.RangeAddress & """. Workbook = """ & ActiveWorkbook.Name & """, ActiveSheet = """ & ActiveSheet.Name & """"
        Case Else
            AppendLog "Unknown error in " & TypeName(Me) & ".Refresh() in section = """ & section_id & """"
    End Select
End Function

Public Function Copy() As TemplateObject
' Action: Makes a copy of itself
'
    Dim output_object As New TemplateObject
    With output_object
        .Log = TypeName(Me) & " object copied" & vbNewLine
        .RangeAddress = Me.RangeAddress
        .Text = Me.Text
        .TextWithWildcards = Me.TextWithWildcards
        .Arguments = Me.Arguments
    End With
        
    Set Copy = output_object
        
End Function

Private Sub AppendLog(NewText As String, Optional NumberOfNewLines As Integer = 1)
' Action: Appends the object log with NewText
'
    pLog = pLog & NewText & String(NumberOfNewLines, vbNewLine)
End Sub

