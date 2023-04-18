VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookRangeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CellAddresses As New Collection
Private m_ColumnCount As Variant
Private m_Formulas As New Collection
Private m_FormulasLocal As New Collection
Private m_FormulasR1C1 As New Collection
Private m_Index As Variant
Private m_NumberFormat As New Collection
Private m_RowCount As Variant
Private m_Text As New Collection
Private m_ValueTypes As New Collection
Private m_Values As New Collection

Public Property Get CellAddresses() As Collection
    Set CellAddresses = m_CellAddresses
End Property

Public Property Set CellAddresses(Value As Collection)
    Set m_CellAddresses = Value
End Property

Public Property Get ColumnCount() As Variant
    ColumnCount = m_ColumnCount
End Property

Public Property Let ColumnCount(Value As Variant)
    m_ColumnCount = Value
End Property

Public Property Get Formulas() As Collection
    Set Formulas = m_Formulas
End Property

Public Property Set Formulas(Value As Collection)
    Set m_Formulas = Value
End Property

Public Property Get FormulasLocal() As Collection
    Set FormulasLocal = m_FormulasLocal
End Property

Public Property Set FormulasLocal(Value As Collection)
    Set m_FormulasLocal = Value
End Property

Public Property Get FormulasR1C1() As Collection
    Set FormulasR1C1 = m_FormulasR1C1
End Property

Public Property Set FormulasR1C1(Value As Collection)
    Set m_FormulasR1C1 = Value
End Property

Public Property Get Index() As Variant
    Index = m_Index
End Property

Public Property Let Index(Value As Variant)
    m_Index = Value
End Property

Public Property Get NumberFormat() As Collection
    Set NumberFormat = m_NumberFormat
End Property

Public Property Set NumberFormat(Value As Collection)
    Set m_NumberFormat = Value
End Property

Public Property Get RowCount() As Variant
    RowCount = m_RowCount
End Property

Public Property Let RowCount(Value As Variant)
    m_RowCount = Value
End Property

Public Property Get Text() As Collection
    Set Text = m_Text
End Property

Public Property Set Text(Value As Collection)
    Set m_Text = Value
End Property

Public Property Get ValueTypes() As Collection
    Set ValueTypes = m_ValueTypes
End Property

Public Property Set ValueTypes(Value As Collection)
    Set m_ValueTypes = Value
End Property

Public Property Get Values() As Collection
    Set Values = m_Values
End Property

Public Property Set Values(Value As Collection)
    Set m_Values = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If CellAddresses.Count > 0 Then .Add "cellAddresses", CellAddresses
        If IsEmpty(ColumnCount) = False Then .Add "columnCount", ColumnCount
        If Formulas.Count > 0 Then .Add "formulas", Formulas
        If FormulasLocal.Count > 0 Then .Add "formulasLocal", FormulasLocal
        If FormulasR1C1.Count > 0 Then .Add "formulasR1C1", FormulasR1C1
        If IsEmpty(Index) = False Then .Add "index", Index
        If NumberFormat.Count > 0 Then .Add "numberFormat", NumberFormat
        If IsEmpty(RowCount) = False Then .Add "rowCount", RowCount
        If Text.Count > 0 Then .Add "text", Text
        If ValueTypes.Count > 0 Then .Add "valueTypes", ValueTypes
        If Values.Count > 0 Then .Add "values", Values
    End With
    Set ToJson = objJson
End Function


