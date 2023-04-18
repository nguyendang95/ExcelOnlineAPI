VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetResource"
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

Public Function GetWorksheet(ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As WorkbookWorksheet
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
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
            Set GetWorksheet = FetchWorksheet(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function AddWorksheet(ByVal WorkbookWorksheet As WorkbookWorksheet, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookWorksheet
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/worksheets"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/worksheets"
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
        .Send JsonConverter.ConvertToJson(WorkbookWorksheet.ToJson)
        .WaitForResponse
        If .Status = 201 Then
            Set AddWorksheet = FetchWorksheet(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function UpdateWorksheet(ByVal WorksheetNameOrId As String, ByVal WorkbookWorksheet As WorkbookWorksheet, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookWorksheet
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
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
        .Send JsonConverter.ConvertToJson(WorkbookWorksheet.ToJson)
        .WaitForResponse
        If .Status = 200 Then
            Set UpdateWorksheet = FetchWorksheet(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub DeleteWorksheet(ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId
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
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Function GetCell(ByVal WorksheetNameOrId As String, ByVal RowIndex As Long, ByVal ColumnIndex As Long, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/cell(row=" & RowIndex & ",column=" & ColumnIndex & ")"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/cell(row=" & RowIndex & ",column=" & ColumnIndex & ")"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/cell(row=" & RowIndex & ",column=" & ColumnIndex & ")"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/cell(row=" & RowIndex & ",column=" & ColumnIndex & ")"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/cell(row=" & RowIndex & ",column=" & ColumnIndex & ")"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/cell(row=" & RowIndex & ",column=" & ColumnIndex & ")"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/cell(row=" & RowIndex & ",column=" & ColumnIndex & ")"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/cell(row=" & RowIndex & ",column=" & ColumnIndex & ")"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/cell(row=" & RowIndex & ",column=" & ColumnIndex & ")"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/cell(row=" & RowIndex & ",column=" & ColumnIndex & ")"
            End If
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
            Set GetCell = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetRange(ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal Address As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/range"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/range"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/range"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/range"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/range"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/range"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/range"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/range"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/range"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/range"
            End If
    End Select
    If Address <> vbNullString Then strUrl = strUrl & "(address='" & Address & "')"
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

Public Function ListTables(ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/tables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/tables"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/tables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/tables"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/tables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/tables"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/tables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/tables"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/tables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/tables"
            End If
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
            Set ListTables = FetchTables(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function AddTable(ByVal Address As String, ByVal HasHeader As Boolean, ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookTable
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/tables/add"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/tables/add"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/tables/add"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/tables/add"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/tables/add"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/tables/add"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/tables/add"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/tables/add"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/tables/add"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/tables/add"
            End If
    End Select
    Dim objRequestBody As Scripting.Dictionary
    Set objRequestBody = New Scripting.Dictionary
    With objRequestBody
        .Add "address", Address
        .Add "hasHeader", HasHeader
    End With
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(objRequestBody)
        .WaitForResponse
        If .Status = 201 Then
            Set AddTable = FetchTable(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListCharts(ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            End If
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
            Set ListCharts = FetchCharts(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function AddChart(ByVal WorkbookChart As WorkbookChart, ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/charts"
            End If
    End Select
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
        If .Status = 201 Then
            Set AddChart = FetchChart(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Function FetchChart(ByVal ResponseText As String) As WorkbookChart
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookChart As WorkbookChart
    Set objWorkbookChart = New WorkbookChart
    With objWorkbookChart
        .Height = objJson.Item("height")
        .Id = objJson.Item("id")
        .Left = objJson.Item("left")
        .Name = objJson.Item("name")
        .Top = objJson.Item("top")
        .Width = objJson.Item("width")
    End With
    Set FetchChart = objWorkbookChart
End Function

Private Function FetchCharts(ByVal ResponseText As String) As Collection
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookChart As WorkbookChart
    Dim colWorkbookCharts As New Collection
    Dim i As Long
    If objJson.Item("value").Count > 0 Then
        For i = 1 To objJson.Item("value").Count
            Set objWorkbookChart = New WorkbookChart
            With objWorkbookChart
                .Height = objJson.Item("value")(i)("height")
                .Id = objJson.Item("value")(i)("id")
                .Left = objJson.Item("value")(i)("left")
                .Name = objJson.Item("value")(i)("name")
                .Top = objJson.Item("value")(i)("top")
                .Width = objJson.Item("value")(i)("width")
            End With
        Next
        colWorkbookCharts.Add objWorkbookChart
    End If
    Set FetchCharts = colWorkbookCharts
End Function

Private Function FetchTable(ByVal ResponseText As String) As WorkbookTable
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookTable As WorkbookTable
    Set objWorkbookTable = New WorkbookTable
    With objWorkbookTable
        .Id = objJson.Item("id")
        .Name = objJson.Item("name")
        .ShowHeaders = objJson.Item("showHeaders")
        .ShowTotals = objJson.Item("showTotals")
        .Style = objJson.Item("style")
        .HighlightFirstColumn = objJson.Item("highlightFirstColumn")
        .HighlightLastColumn = objJson.Item("highlightLastColumn")
        .ShowBandedColumns = objJson.Item("showBandedColumns")
        .ShowBandedRows = objJson.Item("showBandedRows")
        .ShowFilteredButton = objJson.Item("showFilteredButton")
        .LegacyId = objJson.Item("legacyId")
    End With
    Set FetchTable = objWorkbookTable
End Function

Private Function FetchTables(ByVal ResponseText As String) As Collection
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookTable As WorkbookTable
    Dim colWorkbookTables As New Collection
    Dim i As Long
    If objJson.Item("value").Count > 0 Then
        For i = 1 To objJson.Item("value").Count
            Set objWorkbookTable = New WorkbookTable
            With objWorkbookTable
                .Id = objJson.Item("value")(i)("id")
                .Name = objJson.Item("value")(i)("name")
                .ShowHeaders = objJson.Item("value")(i)("showHeaders")
                .ShowTotals = objJson.Item("value")(i)("showTotals")
                .Style = objJson.Item("value")(i)("style")
                .HighlightFirstColumn = objJson.Item("value")(i)("highlightFirstColumn")
                .HighlightLastColumn = objJson.Item("value")(i)("highlightLastColumn")
                .ShowBandedColumns = objJson.Item("value")(i)("showBandedColumns")
                .ShowBandedRows = objJson.Item("value")(i)("showBandedRows")
                .ShowFilteredButton = objJson.Item("value")(i)("showFilteredButton")
                .LegacyId = objJson.Item("value")(i)("legacyId")
            End With
            colWorkbookTables.Add objWorkbookTable
        Next
    End If
    Set FetchTables = colWorkbookTables
End Function

Public Function ListWorkbookNamedItems(ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/names"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/names"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/names"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/names"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/names"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/names"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/names"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/names"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/names"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/names"
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
            Set ListWorkbookNamedItems = FetchWorkbookNamedItems(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetUsedRange(ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal ValuesOnly As Boolean = False, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/usedRange"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/usedRange"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/usedRange"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/usedRange"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/usedRange"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/usedRange"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/usedRange"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/usedRange"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/usedRange"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/usedRange"
            End If
    End Select
    If ValuesOnly Then strUrl = strUrl & "(valuesOnly=true)"
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
            Set GetUsedRange = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListWorksheets(ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/worksheets"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/worksheets"
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
            Set ListWorksheets = FetchWorksheets(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListPivotTables(ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/pivotTables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/pivotTables"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/pivotTables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/pivotTables"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/pivotTables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/pivotTables"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/pivotTables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/pivotTables"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetNameOrId & "/pivotTables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetNameOrId & "/pivotTables"
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
            Set ListPivotTables = FetchPivotTables(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Function FetchPivotTables(ByVal ResponseText As String) As Collection
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookPivotTable As WorkbookPivotTable
    Dim colWorkbookPivotTables As New Collection
    Dim i As Long
    If objJson.Item("value").Count > 0 Then
        For i = 1 To objJson.Item("value").Count
            Set objWorkbookPivotTable = New WorkbookPivotTable
            With objWorkbookPivotTable
                .Name = objJson.Item("value")(i)("name")
                .Id = objJson.Item("value")(i)("id")
            End With
            colWorkbookPivotTables.Add objWorkbookPivotTable
        Next
    End If
    Set FetchPivotTables = colWorkbookPivotTables
End Function

Private Function FetchWorkbookNamedItems(ByVal ResponseText As String) As Collection
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookNamedItem As WorkbookNamedItem
    Dim colWorkbookNamedItems As New Collection
    Dim i As Long
    If objJson.Item("value").Count > 0 Then
        For i = 1 To objJson.Item("value").Count
            Set objWorkbookNamedItem = New WorkbookNamedItem
            With objWorkbookNamedItem
                .Name = objJson.Item("value")(i)("name")
                .Scope = objJson.Item("value")(i)("scope")
                .Comment = objJson.Item("value")(i)("comment")
                .NamedItemType = objJson.Item("value")(i)("type")
                Set .Value = objJson.Item("value")(i)("value")
                .Visible = objJson.Item("value")(i)("visible")
            End With
            colWorkbookNamedItems.Add objWorkbookNamedItem
        Next
    End If
    Set FetchWorkbookNamedItems = colWorkbookNamedItems
End Function

Private Function FetchWorksheet(ByVal ResponseText As String) As WorkbookWorksheet
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookWorksheet As WorkbookWorksheet
    Set objWorkbookWorksheet = New WorkbookWorksheet
    With objWorkbookWorksheet
        .Id = objJson.Item("id")
        .Name = objJson.Item("name")
        .Position = objJson.Item("position")
        .Visibility = objJson.Item("visibility")
    End With
    Set FetchWorksheet = objWorkbookWorksheet
End Function

Private Function FetchWorksheets(ByVal ResponseText As String) As Collection
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookWorksheet As WorkbookWorksheet
    Dim colWorkbookWorksheets As New Collection
    Dim i As Long
    If objJson.Item("value").Count > 0 Then
        For i = 1 To objJson.Item("value").Count
            Set objWorkbookWorksheet = New WorkbookWorksheet
            With objWorkbookWorksheet
                .Id = objJson.Item("value")(i)("id")
                .Name = objJson.Item("value")(i)("name")
                .Position = objJson.Item("value")(i)("position")
                .Visibility = objJson.Item("value")(i)("visibility")
            End With
            colWorkbookWorksheets.Add objWorkbookWorksheet
        Next
    End If
    Set FetchWorksheets = colWorkbookWorksheets
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

