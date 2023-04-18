VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookSessionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_PersistChanges As Variant

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get PersistChanges() As Variant
    PersistChanges = m_PersistChanges
End Property

Public Property Let PersistChanges(Value As Variant)
    m_PersistChanges = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Id <> vbNullString Then .Add "id", Id
        If IsEmpty(PersistChanges) = False Then .Add "persistChanges", PersistChanges
    End With
    Set ToJson = objJson
End Function


