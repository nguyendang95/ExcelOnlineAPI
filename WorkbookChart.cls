VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Height As Variant
Private m_Id As String
Private m_Left As Variant
Private m_Name As String
Private m_Top As Variant
Private m_Width As Variant

Public Property Get Height() As Variant
    Height = m_Height
End Property

Public Property Let Height(Value As Variant)
    m_Height = Value
End Property

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get Left() As Variant
    Left = m_Left
End Property

Public Property Let Left(Value As Variant)
    m_Left = Value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get Top() As Variant
    Top = m_Top
End Property

Public Property Let Top(Value As Variant)
    m_Top = Value
End Property

Public Property Get Width() As Variant
    Width = m_Width
End Property

Public Property Let Width(Value As Variant)
    m_Width = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(Height) = False Then .Add "height", Height
        If Id <> vbNullString Then .Add "id", Id
        If IsEmpty(Left) = False Then .Add "left", Left
        If Name <> vbNullString Then .Add "name", Name
        If IsEmpty(Top) = False Then .Add "top", Top
        If IsEmpty(Width) = False Then .Add "width", Width
    End With
    Set ToJson = objJson
End Function



