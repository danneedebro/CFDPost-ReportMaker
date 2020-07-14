VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlotObjectCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES
Private pCollection As Collection
'
'
' PROPERTIES
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
' Action: Enables for each loops (VB_UserMemId = -4)
    Set NewEnum = pCollection.[_NewEnum]
End Property
'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set pCollection = New Collection
End Sub

Public Sub Add(ByVal Item As PlotObject, Optional Key As String = "")
    If Key = "" Then
        pCollection.Add Item
    Else
        pCollection.Add Item, Key:=Key
    End If
End Sub

Public Function Item(index As Variant) As PlotObject
Attribute Item.VB_UserMemId = 0
    Set Item = pCollection(index)
End Function

Public Function Count() As Long
    Count = pCollection.Count
End Function

Public Function FindItemByName(Name As String) As PlotObject
' Action: Return the plot object defined by a name
'
    Dim plot_object As PlotObject
    For Each plot_object In pCollection
        If plot_object.Name = Name Then
            Set FindItemByName = plot_object
            Exit For
        End If
    Next plot_object
    
    
End Function

'
'
' CLASS METHODS

'
'
' STATIC METHODS

