VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookFilterCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Color As String
Private m_Criterion1 As String
Private m_Criterion2 As String
Private m_DynamicCriteria As String
Private m_FilterOn As String
Private m_Values As New Collection
Private m_Icon As WorkbookIcon

Public Property Get Color() As String
    Color = m_Color
End Property

Public Property Let Color(Value As String)
    m_Color = Value
End Property

Public Property Get Criterion1() As String
    Criterion1 = m_Criterion1
End Property

Public Property Let Criterion1(Value As String)
    m_Criterion1 = Value
End Property

Public Property Get Criterion2() As String
    Criterion2 = m_Criterion2
End Property

Public Property Let Criterion2(Value As String)
    m_Criterion2 = Value
End Property

Public Property Get DynamicCriteria() As String
    DynamicCriteria = m_DynamicCriteria
End Property

Public Property Let DynamicCriteria(Value As String)
    m_DynamicCriteria = Value
End Property

Public Property Get FilterOn() As String
    FilterOn = m_FilterOn
End Property

Public Property Let FilterOn(Value As String)
    m_FilterOn = Value
End Property

Public Property Get Values() As Collection
    Set Values = m_Values
End Property

Public Property Set Values(Value As Collection)
    Set m_Values = Value
End Property

Public Property Get Icon() As WorkbookIcon
    Set Icon = m_Icon
End Property

Public Property Set Icon(Value As WorkbookIcon)
    Set m_Icon = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Color <> vbNullString Then .Add "color", Color
        If Criterion1 <> vbNullString Then .Add "criterion1", Criterion1
        If Criterion2 <> vbNullString Then .Add "criterion2", Criterion2
        If DynamicCriteria <> vbNullString Then .Add "dynamicCriteria", DynamicCriteria
        If FilterOn <> vbNullString Then .Add "filterOn", FilterOn
        If Values.Count > 0 Then .Add "values", Values
        If Not Icon Is Nothing Then .Add "icon", Icon.ToJson
    End With
    Set ToJson = objJson
End Function

