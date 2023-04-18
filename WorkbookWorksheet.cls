VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_Name As String
Private m_Position As Variant
Private m_Visibility As String

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get Position() As Variant
    Position = m_Position
End Property

Public Property Let Position(Value As Variant)
    m_Position = Value
End Property

Public Property Get Visibility() As String
    Visibility = m_Visibility
End Property

Public Property Let Visibility(Value As String)
    m_Visibility = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Id <> vbNullString Then .Add "id", Id
        If Name <> vbNullString Then .Add "name", Name
        If IsEmpty(Position) = False Then .Add "position", Position
        If Visibility <> vbNullString Then .Add "visibility", Visibility
    End With
    Set ToJson = objJson
End Function


