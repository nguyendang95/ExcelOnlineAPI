VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CalculationMode As String

Public Property Get CalculationMode() As String
    CalculationMode = m_CalculationMode
End Property

Public Property Let CalculationMode(Value As String)
    m_CalculationMode = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If CalculationMode <> vbNullString Then .Add "calculationMode", CalculationMode
    End With
    Set ToJson = objJson
End Function


