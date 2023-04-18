VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookNamedItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Name As String
Private m_Comment As String
Private m_Scope As String
Private m_Type As String
Private m_Value As New Collection
Private m_Visible As Variant

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Public Property Let Comment(Value As String)
    m_Comment = Value
End Property

Public Property Get Scope() As String
    Scope = m_Scope
End Property

Public Property Let Scope(Value As String)
    m_Scope = Value
End Property

Public Property Get NamedItemType() As String
    NamedItemType = m_Type
End Property

Public Property Let NamedItemType(Value As String)
    m_Type = Value
End Property

Public Property Get Value() As Collection
    Set Value = m_Value
End Property

Public Property Set Value(Value As Collection)
    Set m_Value = Value
End Property

Public Property Get Visible() As Variant
    Visible = m_Visible
End Property

Public Property Let Visible(Value As Variant)
    m_Visible = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Name <> vbNullString Then .Add "name", Name
        If Comment <> vbNullString Then .Add "comment", Comment
        If Scope <> vbNullString Then .Add "scope", Scope
        If NamedItemType <> vbNullString Then .Add "type", NamedItemType
        If Value.Count > 0 Then .Add "value", Value
        If IsEmpty(Visible) = False Then .Add "visible", Visible
    End With
    Set ToJson = objJson
End Function

