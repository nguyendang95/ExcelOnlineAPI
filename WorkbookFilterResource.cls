VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookFilterResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_AccessToken As String

Friend Property Get AccessToken() As String
   AccessToken = m_AccessToken
End Property

Friend Property Let AccessToken(Value As String)
   m_AccessToken = Value
End Property

Public Function Apply(GetTableColumn As GetTableColumn, ByVal Criteria As WorkbookFilterCriteria, ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal ColumnIndex As Long, Optional ByVal ColumnName As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookFilterCriteria
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            End If
    End Select
    Select Case GetTableColumn
        Case ByIndex
            strUrl = strUrl & "/columns/itemAt(index=" & ColumnIndex & ")"
        Case ByName
            strUrl = strUrl & "/columms/" & objWebUtilities.URLEncode(ColumnName)
    End Select
    strUrl = strUrl & "/filter/apply"
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(Criteria.ToJson)
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Function

Public Function Clear(GetTableColumn As GetTableColumn, ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal ColumnIndex As Long, Optional ByVal ColumnName As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookFilterCriteria
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/" & TableIdOrName
                End Select
            End If
    End Select
    Select Case GetTableColumn
        Case ByIndex
            strUrl = strUrl & "/columns/itemAt(index=" & ColumnIndex & ")"
        Case ByName
            strUrl = strUrl & "/columms/" & objWebUtilities.URLEncode(ColumnName)
    End Select
    strUrl = strUrl & "/filter/clear"
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Function

Private Function FetchWorkbookFilter(ByVal ResponseText As String) As WorkbookFilter
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objFilter As WorkbookFilter
    Set objFilter = New WorkbookFilter
    Dim objFilterCriteria As WorkbookFilterCriteria
    Set objFilterCriteria = New WorkbookFilterCriteria
    If objJson.Exists("criteria") Then
        With objFilterCriteria
            .Color = objJson.Item("criteria")("color")
            .Criterion1 = objJson.Item("criteria")("criterion1")
            .Criterion2 = objJson.Item("criteria")("criterion2")
            .DynamicCriteria = objJson.Item("criteria")("dynamicCriteria")
            .FilterOn = objJson.Item("criteria")("filterOn")
            Set .Values = objJson.Item("criteria")("values")
        End With
    End If
    Set objFilter.Criteria = objFilterCriteria
    Set FetchWorkbookFilter = objFilter
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub

