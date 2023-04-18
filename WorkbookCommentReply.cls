VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookCommentReply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Content As String
Private m_ContentType As String
Private m_Id As String

Public Property Get Content() As String
    Content = m_Content
End Property

Public Property Let Content(Value As String)
    m_Content = Value
End Property

Public Property Get ContentType() As String
    ContentType = m_ContentType
End Property

Public Property Let ContentType(Value As String)
    m_ContentType = Value
End Property

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Content <> vbNullString Then .Add "content", Content
        If ContentType <> vbNullString Then .Add "contentType", ContentType
        If Id <> vbNullString Then .Add "id", Id
    End With
    Set ToJson = objJson
End Function



