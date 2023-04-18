VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookTableColumnResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum GetTableColumn
    ByIndex
    ByName
End Enum

Private m_AccessToken As String

Friend Property Get AccessToken() As String
   AccessToken = m_AccessToken
End Property

Friend Property Let AccessToken(Value As String)
   m_AccessToken = Value
End Property

Public Function GetColumn(GetTableColumn As GetTableColumn, ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal ColumnIndex As Long, Optional ByVal ColumnName As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As WorkbookTableColumn
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
            strUrl = strUrl & "/columns/" & objWebUtilities.URLEncode(ColumnName)
    End Select
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetColumn = FetchWorkbookTableColumn(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function UpdateColumn(GetTableColumn As GetTableColumn, ByVal Values As Variant, ByVal WorkbookTableColumn As WorkbookTableColumn, ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal ColumnIndex As Long, Optional ByVal ColumnName As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookTableColumn
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
            strUrl = strUrl & "/columns/" & objWebUtilities.URLEncode(ColumnName)
    End Select
    Dim objRequestBody As Scripting.Dictionary
    Set objRequestBody = New Scripting.Dictionary
    objRequestBody.Add "values", JsonConverter.ParseJson(JsonConverter.ConvertToJson(Values))
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "PATCH", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(objRequestBody)
        .WaitForResponse
        If .Status = 200 Then
            Set UpdateColumn = FetchWorkbookTableColumn(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetDataBodyRange(GetTableColumn As GetTableColumn, ByVal WorkbookTableColumn As WorkbookTableColumn, ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal ColumnIndex As Long, Optional ByVal ColumnName As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
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
            strUrl = strUrl & "/columns/itemAt(index=" & ColumnIndex & ")/dataBodyRange"
        Case ByName
            strUrl = strUrl & "/columns/" & objWebUtilities.URLEncode(ColumnName) & "/dataBodyRange"
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetDataBodyRange = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetHeaderRowRange(ByVal WorkbookTableColumn As WorkbookTableColumn, ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal ColumnIndex As Long, Optional ByVal ColumnName As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
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

    strUrl = strUrl & "/columns/itemAt(index=" & ColumnIndex & ")/headerRowRange"
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetHeaderRowRange = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetRange(GetTableColumn As GetTableColumn, ByVal WorkbookTableColumn As WorkbookTableColumn, ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal ColumnIndex As Long, Optional ByVal ColumnName As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
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
            strUrl = strUrl & "/columns/itemAt(index=" & ColumnIndex & ")/range"
        Case ByName
            strUrl = strUrl & "/columns/" & objWebUtilities.URLEncode(ColumnName) & "/range"
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetRange = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetTotalRowRange(GetTableColumn As GetTableColumn, ByVal WorkbookTableColumn As WorkbookTableColumn, ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal ColumnIndex As Long, Optional ByVal ColumnName As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
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
            strUrl = strUrl & "/columns/itemAt(index=" & ColumnIndex & ")/totalRowRange"
        Case ByName
            strUrl = strUrl & "/columns/" & objWebUtilities.URLEncode(ColumnName) & "/totalRowRange"
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetTotalRowRange = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub DeleteColumn(GetTableColumn As GetTableColumn, ByVal WorkbookTableColumn As WorkbookTableColumn, ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal ColumnIndex As Long, Optional ByVal ColumnName As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
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
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "DELETE", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Function ListColumns(GetTableColumn As GetTableColumn, ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As Collection
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
    strUrl = strUrl & "/columns"
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set ListColumns = FetchWorkbookTableColumns(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function CreateColumn(ByVal WorkbookTableColumn As WorkbookTableColumn, ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookTableColumn
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
    strUrl = strUrl & "/columns"
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(WorkbookTableColumn.ToJson)
        .WaitForResponse
        If .Status = 201 Then
            Set CreateColumn = FetchWorkbookTableColumn(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Function FetchWorkbookTableColumn(ByVal ResponseText As String) As WorkbookTableColumn
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookTableColumn As WorkbookTableColumn
    Set objWorkbookTableColumn = New WorkbookTableColumn
    With objWorkbookTableColumn
        .Index = objJson.Item("index")
        .Name = objJson.Item("name")
        .Id = objJson.Item("id")
        Set .Values = objJson.Item("values")
    End With
    Set FetchWorkbookTableColumn = objWorkbookTableColumn
End Function

Private Function FetchWorkbookTableColumns(ByVal ResponseText As String) As Collection
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim i As Long
    Dim objWorkbookTableColumn As WorkbookTableColumn
    Dim colWorkbookTableColumns As New Collection
    Set objWorkbookTableColumn = New WorkbookTableColumn
    If objJson.Item("value").Count > 0 Then
        For i = 1 To objJson.Item("value").Count
            With objWorkbookTableColumn
                .Index = objJson.Item("value")(i)("index")
                .Name = objJson.Item("value")(i)("name")
                .Id = objJson.Item("value")(i)("id")
                Set .Values = objJson.Item("value")(i)("values")
            End With
            colWorkbookTableColumns.Add objWorkbookTableColumn
        Next
    End If
    Set FetchWorkbookTableColumns = colWorkbookTableColumns
End Function

Private Function FetchRange(ByVal ResponseText As String) As WorkbookRange
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookRange As WorkbookRange
    Set objWorkbookRange = New WorkbookRange
    With objWorkbookRange
        .Address = objJson.Item("address")
        .AddressLocal = objJson.Item("addressLocal")
        .ColumnCount = objJson.Item("columnCount")
        .CellCount = objJson.Item("cellCount")
        .ColumnHidden = objJson.Item("columnHidden")
        .ColumnIndex = objJson.Item("columnIndex")
        .RowHidden = objJson.Item("rowHidden")
        .RowIndex = objJson.Item("rowIndex")
        .Hidden = objJson.Item("hidden")
        Set .ValueTypes = objJson.Item("valueTypes")
        Set .Formulas = objJson.Item("formulas")
        Set .FormulasLocal = objJson.Item("formulasLocal")
        Set .FormulasR1C1 = objJson.Item("formulasR1C1")
        Set .Text = objJson.Item("text")
        Set .Values = objJson.Item("values")
    End With
    Set FetchRange = objWorkbookRange
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
