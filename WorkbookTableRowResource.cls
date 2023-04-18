VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookTableRowResource"
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

Public Function GetItemAt(Location As Location, ByVal WorkbookTableIdOrName As String, ByVal WorkbookTableRow As WorkbookTableRow, Optional ByVal WorksheetId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal WorkbookSessionId As String) As WorkbookTableRow
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rowsitemAt"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(WorkbookTableRow.ToJson)
        .WaitForResponse
        If .Status = 200 Then
            Set GetItemAt = FetchWorkbookTableRow(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetWorkbookTableRow(Location As Location, ByVal WorkbookTableIdOrName As String, ByVal RowIndex As Long, Optional ByVal WorksheetId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ODataQuery As String) As WorkbookTableRow
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rowsitemAt(index=" & RowIndex & ")"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
            End If
    End Select
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetWorkbookTableRow = FetchWorkbookTableRow(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListWorkbookTableRows(Location As Location, ByVal WorkbookTableIdOrName As String, Optional ByVal WorksheetId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ODataQuery As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows"
                Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows"
                Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows"
                Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows"
                Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows"
                Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows"
            End If
    End Select
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set ListWorkbookTableRows = FetchWorkbookTableRows(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function UpdateRow(Location As Location, ByVal WorkbookTableIdOrName As String, ByVal RowIndex As Long, ByVal Values As Scripting.Dictionary, Optional ByVal WorksheetId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal WorkbookSessionId As String) As WorkbookTableRow
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "PATCH", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(Values)
        .WaitForResponse
        If .Status = 200 Then
            Set UpdateRow = FetchWorkbookTableRow(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub DeleteWorkbookTableRow(Location As Location, ByVal WorkbookTableIdOrName As String, ByVal RowIndex As Long, Optional ByVal WorksheetId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal WorkbookSessionId As String)
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "DELETE", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status <> 200 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Function GetRange(Location As Location, ByVal WorkbookTableIdOrName As String, ByVal RowIndex As Long, Optional ByVal WorksheetId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal WorkbookSessionId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")/range"
                Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")/range"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rowsitemAt(index=" & RowIndex & ")/range"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")/range"
                Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")/range"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")/range"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")/range"
                Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")/range"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")/range"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")/range"
                Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")/range"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")/range"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")/range"
                Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")/range"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt(index=" & RowIndex & ")/range"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
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

Public Function AddRow(Location As Location, ByVal WorkbookTableIdOrName As String, ByVal WorkbookTableRow As WorkbookTableRow, Optional ByVal WorksheetId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal WorkbookSessionId As String) As WorkbookTableRow
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rowsitemAt"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                If WorksheetId <> vbNullString Then
                    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetId & "/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
                End If
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & ":/workbook/tables/" & WorkbookTableIdOrName & "/rows/itemAt"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(WorkbookTableRow.ToJson)
        .WaitForResponse
        If .Status = 201 Then
            Set AddRow = FetchWorkbookTableRow(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
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

Private Function FetchWorkbookTableRow(ByVal ResponseText As String) As WorkbookTableRow
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim i As Long
    Dim objWorkbookTableRow As WorkbookTableRow
    Dim colWorkbookTableRows As New Collection
    Set objWorkbookTableRow = New WorkbookTableRow
    With objWorkbookTableRow
        .Index = objJson.Item("index")
        Set .Values = objJson.Item("values")
    End With
    Set FetchWorkbookTableRow = objWorkbookTableRow
End Function

Private Function FetchWorkbookTableRows(ByVal ResponseText As String) As Collection
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim i As Long
    Dim objWorkbookTableRow As WorkbookTableRow
    Dim colWorkbookTableRows As New Collection
    If objJson.Item("value").Count > 0 Then
        For i = 1 To objJson.Item("value").Count
            Set objWorkbookTableRow = New WorkbookTableRow
            With objWorkbookTableRow
                .Index = objJson.Item("value")(i)("index")
                Set .Values = objJson.Item("value")(i)("values")
            End With
            colWorkbookTableRows.Add objWorkbookTableRow
        Next
    End If
    Set FetchWorkbookTableRows = colWorkbookTableRows
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
