VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Criteria As WorkbookFilterCriteria

Public Property Get Criteria() As WorkbookFilterCriteria
    Set Criteria = m_Criteria
End Property

Public Property Set Criteria(Value As WorkbookFilterCriteria)
    Set m_Criteria = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Not Criteria Is Nothing Then .Add "criteria", Criteria.ToJson
    End With
    Set ToJson = objJson
End Function



