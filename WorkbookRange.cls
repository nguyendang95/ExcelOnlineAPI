VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Address As String
Private m_AddressLocal As String
Private m_CellCount As Variant
Private m_ColumnCount As Variant
Private m_ColumnHidden As Variant
Private m_ColumnIndex As Variant
Private m_Formulas As New Collection
Private m_FormulasLocal As New Collection
Private m_FormulasR1C1 As New Collection
Private m_Hidden As Variant
Private m_NumberFormat As Variant
Private m_RowCount As Variant
Private m_RowHidden As Variant
Private m_RowIndex As Variant
Private m_Text As New Collection
Private m_ValueTypes As New Collection
Private m_Values As New Collection

Public Property Get Address() As String
    Address = m_Address
End Property

Public Property Let Address(Value As String)
    m_Address = Value
End Property

Public Property Get AddressLocal() As String
    AddressLocal = m_AddressLocal
End Property

Public Property Let AddressLocal(Value As String)
    m_AddressLocal = Value
End Property

Public Property Get CellCount() As Variant
    CellCount = m_CellCount
End Property

Public Property Let CellCount(Value As Variant)
    m_CellCount = Value
End Property

Public Property Get ColumnCount() As Variant
    ColumnCount = m_ColumnCount
End Property

Public Property Let ColumnCount(Value As Variant)
    m_ColumnCount = Value
End Property

Public Property Get ColumnHidden() As Variant
    ColumnHidden = m_ColumnHidden
End Property

Public Property Let ColumnHidden(Value As Variant)
    m_ColumnHidden = Value
End Property

Public Property Get ColumnIndex() As Variant
    ColumnIndex = m_ColumnIndex
End Property

Public Property Let ColumnIndex(Value As Variant)
    m_ColumnIndex = Value
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

Public Property Get Hidden() As Variant
    Hidden = m_Hidden
End Property

Public Property Let Hidden(Value As Variant)
    m_Hidden = Value
End Property

Public Property Get NumberFormat() As Variant
    NumberFormat = m_NumberFormat
End Property

Public Property Let NumberFormat(Value As Variant)
    m_NumberFormat = Value
End Property

Public Property Get RowCount() As Variant
    RowCount = m_RowCount
End Property

Public Property Let RowCount(Value As Variant)
    m_RowCount = Value
End Property

Public Property Get RowHidden() As Variant
    RowHidden = m_RowHidden
End Property

Public Property Let RowHidden(Value As Variant)
    m_RowHidden = Value
End Property

Public Property Get RowIndex() As Variant
    RowIndex = m_RowIndex
End Property

Public Property Let RowIndex(Value As Variant)
    m_RowIndex = Value
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
        If Address <> vbNullString Then .Add "address", Address
        If AddressLocal <> vbNullString Then .Add "addressLocal", AddressLocal
        If IsEmpty(CellCount) = False Then .Add "cellCount", CellCount
        If IsEmpty(ColumnCount) = False Then .Add "columnCount", ColumnCount
        If IsEmpty(ColumnHidden) = False Then .Add "columnHidden", ColumnHidden
        If IsEmpty(ColumnIndex) = False Then .Add "columnIndex", ColumnIndex
        If Formulas.Count > 0 Then .Add "formulas", Formulas
        If FormulasLocal.Count > 0 Then .Add "formulasLocal", FormulasLocal
        If FormulasR1C1.Count > 0 Then .Add "formulasR1C1", FormulasR1C1
        If IsEmpty(Hidden) = False Then .Add "hidden", Hidden
        If IsEmpty(NumberFormat) = False Then .Add "numberFormat", NumberFormat
        If IsEmpty(RowCount) = False Then .Add "rowCount", RowCount
        If IsEmpty(RowHidden) = False Then .Add "rowHidden", RowHidden
        If IsEmpty(RowIndex) = False Then .Add "rowIndex", RowIndex
        If Text.Count > 0 Then .Add "text", Text
        If ValueTypes.Count > 0 Then .Add "valueTypes", "valueTypes"
        If Values.Count > 0 Then .Add "values", Values
    End With
    Set ToJson = objJson
End Function

