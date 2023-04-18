VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookTableColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_Index As Variant
Private m_Name As String
Private m_Values As New Collection

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get Index() As Variant
    Index = m_Index
End Property

Public Property Let Index(Value As Variant)
    m_Index = Value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get Values() As Collection
    Set Values = m_Values
End Property

Public Property Set Values(Value As Collection)
    Set m_Values = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Id <> vbNullString Then .Add "id", Id
        If IsEmpty(Index) = False Then .Add "index", Index
        If Name <> vbNullString Then .Add "name", Name
        If Values.Count > 0 Then .Add "values", Values
    End With
    Set ToJson = objJson
End Function


