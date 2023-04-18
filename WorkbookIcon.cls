VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_SetValue As String
Private m_Index As Variant

Public Property Get SetValue() As String
    SetValue = m_SetValue
End Property

Public Property Let SetValue(Value As String)
    m_SetValue = Value
End Property

Public Property Get Index() As Variant
    Index = m_Index
End Property

Public Property Let Index(Value As Variant)
    m_Index = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If SetValue <> vbNullString Then .Add "set", SetValue
        If IsEmpty(Index) = False Then .Add "index", Index
    End With
    Set ToJson = objJson
End Function



