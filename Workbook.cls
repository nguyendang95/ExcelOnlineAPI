VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Workbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Names As New Collection
Private m_Tables As New Collection
Private m_Worksheets As New Collection

Public Property Get Names() As Collection
    Set Names = m_Names
End Property

Public Property Set Names(Value As Collection)
    Set m_Names = Value
End Property

Public Property Get Tables() As Collection
    Set Tables = m_Tables
End Property

Public Property Set Tables(Value As Collection)
    Set m_Tables = Value
End Property

Public Property Get Worksheets() As Collection
    Set Worksheets = m_Worksheets
End Property

Public Property Set Worksheets(Value As Collection)
    Set m_Worksheets = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Names.Count > 0 Then .Add "names", Names
        If Tables.Count > 0 Then .Add "tables", Tables
        If Worksheets.Count > 0 Then .Add "worksheets", Worksheets
    End With
    Set ToJson = objJson
End Function



