VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_HighlightFirstColumn As Variant
Private m_HighlightLastColumn As Variant
Private m_Id As String
Private m_Name As String
Private m_ShowBandedColumns As Variant
Private m_ShowBandedRows As Variant
Private m_ShowFilteredButton As Variant
Private m_ShowHeaders As Variant
Private m_ShowTotals As Variant
Private m_Style As String
Private m_LegacyId As String

Public Property Get HighlightFirstColumn() As Variant
    HighlightFirstColumn = m_HighlightFirstColumn
End Property

Public Property Let HighlightFirstColumn(Value As Variant)
    m_HighlightFirstColumn = Value
End Property

Public Property Get HighlightLastColumn() As Variant
    HighlightLastColumn = m_HighlightLastColumn
End Property

Public Property Let HighlightLastColumn(Value As Variant)
    m_HighlightLastColumn = Value
End Property

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

Public Property Get ShowBandedColumns() As Variant
    ShowBandedColumns = m_ShowBandedColumns
End Property

Public Property Let ShowBandedColumns(Value As Variant)
    m_ShowBandedColumns = Value
End Property

Public Property Get ShowBandedRows() As Variant
    ShowBandedRows = m_ShowBandedRows
End Property

Public Property Let ShowBandedRows(Value As Variant)
    m_ShowBandedRows = Value
End Property

Public Property Get ShowFilteredButton() As Variant
    ShowFilteredButton = m_ShowFilteredButton
End Property

Public Property Let ShowFilteredButton(Value As Variant)
    m_ShowFilteredButton = Value
End Property

Public Property Get ShowHeaders() As Variant
    ShowHeaders = m_ShowHeaders
End Property

Public Property Let ShowHeaders(Value As Variant)
    m_ShowHeaders = Value
End Property

Public Property Get ShowTotals() As Variant
    ShowTotals = m_ShowTotals
End Property

Public Property Let ShowTotals(Value As Variant)
    m_ShowTotals = Value
End Property

Public Property Get Style() As String
    Style = m_Style
End Property

Public Property Let Style(Value As String)
    m_Style = Value
End Property

Public Property Get LegacyId() As String
    LegacyId = m_LegacyId
End Property

Public Property Let LegacyId(Value As String)
    m_LegacyId = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(HighlightFirstColumn) = False Then .Add "highlightFirstColumn", HighlightFirstColumn
        If IsEmpty(HighlightLastColumn) = False Then .Add "highlightLastColumn", HighlightLastColumn
        If Id <> vbNullString Then .Add "id", Id
        If Name <> vbNullString Then .Add "name", Name
        If IsEmpty(ShowBandedColumns) = False Then .Add "showBandedColumns", ShowBandedColumns
        If IsEmpty(ShowBandedRows) = False Then .Add "showBandedRows", ShowBandedRows
        If IsEmpty(ShowFilteredButton) = False Then .Add "showFilteredButton", ShowFilteredButton
        If IsEmpty(ShowHeaders) = False Then .Add "showHeaders", ShowHeaders
        If IsEmpty(ShowTotals) = False Then .Add "showTotals", ShowTotals
        If Style <> vbNullString Then .Add "style", Style
        If LegacyId <> vbNullString Then .Add "legacyId", LegacyId
    End With
    Set ToJson = objJson
End Function

