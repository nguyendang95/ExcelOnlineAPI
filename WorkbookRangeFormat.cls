VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookRangeFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ColumnWidth As Variant
Private m_HorizontalAlignment As String
Private m_RowHeight As Variant
Private m_VerticalAlignment As String
Private m_WrapText As Variant

Public Property Get ColumnWidth() As Variant
    ColumnWidth = m_ColumnWidth
End Property

Public Property Let ColumnWidth(Value As Variant)
    m_ColumnWidth = Value
End Property

Public Property Get HorizontalAlignment() As String
    HorizontalAlignment = m_HorizontalAlignment
End Property

Public Property Let HorizontalAlignment(Value As String)
    m_HorizontalAlignment = Value
End Property

Public Property Get RowHeight() As Variant
    RowHeight = m_RowHeight
End Property

Public Property Let RowHeight(Value As Variant)
    m_RowHeight = Value
End Property

Public Property Get VerticalAlignment() As String
    VerticalAlignment = m_VerticalAlignment
End Property

Public Property Let VerticalAlignment(Value As String)
    m_VerticalAlignment = Value
End Property

Public Property Get WrapText() As Variant
    WrapText = m_WrapText
End Property

Public Property Let WrapText(Value As Variant)
    m_WrapText = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(ColumnWidth) = False Then .Add "columnWidth", ColumnWidth
        If HorizontalAlignment <> vbNullString Then .Add "horizontalAlignment", HorizontalAlignment
        If IsEmpty(RowHeight) = False Then .Add "rowHeight", RowHeight
        If VerticalAlignment <> vbNullString Then .Add "verticalAlignment", VerticalAlignment
        If IsEmpty(WrapText) = False Then .Add "wrapText", WrapText
    End With
    Set ToJson = objJson
End Function


