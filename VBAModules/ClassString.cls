VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pString As String

Private Sub Class_Initialize()
    pString = ""
End Sub

Public Sub Clear()
' Action: Clears string
'
    pString = ""
End Sub

Public Sub AppendRow(NewText As String)
' Action: Appends a new row with text
'
    pString = pString & NewText & vbNewLine
End Sub

Public Sub Append(NewText As String)
' Action: Appends a new row with text
'
    pString = pString & NewText
End Sub

Public Function output() As String
' Action: Returns contents
'
    output = pString
End Function

Public Sub NewLines(NumberOfNewLines As Integer)
    pString = pString & String(NumberOfNewLines, vbNewLine)
End Sub
