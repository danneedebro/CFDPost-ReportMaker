VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES
Private pRangeObject As Range
Public Log As String
'
'
' PROPERTIES

'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Me.Log = "Object created" & vbNewLine
End Sub

Public Function Create(TableName As String) As Boolean
    On Error GoTo ErrorHandler
    Dim success As Boolean, section_id As String
    success = True
    
    section_id = "Reset object"
    AppendLog "TableObject.Create(TableName:=""" & TableName & """)"
    
    ThisWorkbook.Activate
    Set pRangeObject = Range(TableName)

    
    Create = success
    Exit Function
ErrorHandler:
    Create = False
    AppendLog "Unknown error in Create function, section_id = """ & section_id & """"
End Function

Private Sub AppendLog(NewText As String, Optional NewLines As Integer = 1)
    Me.Log = Me.Log & NewText & String(NewLines, vbNewLine)
End Sub

Public Function Item() As Range
Attribute Item.VB_UserMemId = 0
    Set Item = pRangeObject
End Function

Public Function GetRow(RowIndex As Long) As Range
    With pRangeObject
        Set GetRow = .Range(.Cells(RowIndex, 1), .Cells(RowIndex, .Columns.Count))
    End With
End Function

Public Function GetRange() As Range
    Set GetRange = pRangeObject
End Function

Public Function NumberOfRows() As Integer
    NumberOfRows = pRangeObject.Rows.Count
End Function

Public Sub SelectRow(RowNumber As Integer, Optional ColNumber As Integer = 0)
    
    Dim range_to_select As Range
    
    With pRangeObject
        Set range_to_select = Range(.Cells(RowNumber, 1), .Cells(RowNumber, .Columns.Count))
        range_to_select.Worksheet.Activate
        range_to_select.Select
    End With
    
End Sub

'
'
' CLASS METHODS

'
'
' STATIC METHODS

