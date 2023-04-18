VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Location
    LocationDrive
    LocationGroup
    LocationSignedInUser
    LocationSite
    LocationOtherUser
    LocationShare
End Enum

Private m_AccessToken As String

Friend Property Get AccessToken() As String
   AccessToken = m_AccessToken
End Property

Friend Property Let AccessToken(Value As String)
   m_AccessToken = Value
End Property

Public Function CreateSessionWithLongRunningOperationPattern(ByVal WorkbookSessionInfo As WorkbookSessionInfo, Location As Location, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As String
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/createSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/functions/createSession"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/createSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/createSession"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/createSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/createSession"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/createSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/createSession"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/createSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/createSession"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "Prefer", "respond-async"
        .Send JsonConverter.ConvertToJson(WorkbookSessionInfo.ToJson)
        .WaitForResponse
        If .Status = 202 Then
            CreateSessionWithLongRunningOperationPattern = .GetResponseHeader("Location")
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function CreateSession(ByVal WorkbookSessionInfo As WorkbookSessionInfo, Location As Location, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookSessionInfo
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/createSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/functions/createSession"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/createSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/createSession"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/createSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/createSession"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/createSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/createSession"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/createSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/createSession"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(WorkbookSessionInfo.ToJson)
        .WaitForResponse
        If .Status = 201 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objWorkbookSessionInfo As WorkbookSessionInfo
            Set objWorkbookSessionInfo = New WorkbookSessionInfo
            With objWorkbookSessionInfo
                .Id = objJson.Item("id")
                .PersistChanges = objJson.Item("persistChanges")
            End With
            Set CreateSession = objWorkbookSessionInfo
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ExecuteWorkbookFunction(ByVal WorkbookFunctionName As String, ByVal Parameters As Scripting.Dictionary, Location As Location, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal WorkbookSessionId As String) As Scripting.Dictionary
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/" & WorkbookFunctionName
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/functions/" & WorkbookFunctionName
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/" & WorkbookFunctionName
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/functions/" & WorkbookFunctionName
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/" & WorkbookFunctionName
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/functions/" & WorkbookFunctionName
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/" & WorkbookFunctionName
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/functions/" & WorkbookFunctionName
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/" & WorkbookFunctionName
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/functions/" & WorkbookFunctionName
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
        .Send JsonConverter.ConvertToJson(Parameters)
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Set ExecuteWorkbookFunction = objJson
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub RefreshSession(ByVal WorkbookSessionId As String, Location As Location, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/refreshSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/refreshSession"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/refreshSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/refreshSession"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/refreshSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/refreshSession"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/refreshSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/refreshSession"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/refreshSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/refreshSession"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Sub CloseSession(ByVal WorkbookSessionId As String, Location As Location, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/closeSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/closeSession"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/closeSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/closeSession"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/closeSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/closeSession"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/closeSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/closeSession"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/closeSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/closeSession"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Function ListTables(Location As Location, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal WorkbookSessionId As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables"
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
            Set ListTables = FetchTables(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListWorksheets(Location As Location, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal WorkbookSessionId As String) As Collection
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

Public Function ListNamedItems(Location As Location, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ODataQuery As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook"
            End If
    End Select
    strUrl = strUrl & "/names"
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
            Set ListNamedItems = FetchNamedItems(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Function FetchNamedItems(ByVal ResponseText As String) As Collection
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookNamedItem As WorkbookNamedItem
    Dim i As Long
    Dim colWorkbookNamedItems As New Collection
    If objJson.Item("value").Count > 0 Then
        For i = 1 To objJson.Item("value").Count
            Set objWorkbookNamedItem = New WorkbookNamedItem
            With objWorkbookNamedItem
                .Name = objJson.Item("value")(i)("name")
                .NamedItemType = objJson.Item("value")(i)("type")
                Set .Value = objJson.Item("value")(i)("value")
                .Visible = objJson.Item("value")(i)("visible")
            End With
            colWorkbookNamedItems.Add objWorkbookNamedItem
        Next
    End If
    Set FetchNamedItems = colWorkbookNamedItems
End Function

Private Function FetchWorksheets(ByVal ResponseText As String) As Collection
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookWorksheet As WorkbookWorksheet
    Dim i As Long
    Dim colWorkbookWorksheets As New Collection
    If objJson.Item("value").Count > 0 Then
        For i = 1 To objJson.Item("value").Count
            Set objWorkbookWorksheet = New WorkbookWorksheet
            With objWorkbookWorksheet
                .Id = objJson.Item("value")(i)("id")
                .Position = objJson.Item("value")(i)("position")
                .Name = objJson.Item("value")(i)("name")
                .Visibility = objJson.Item("value")(i)("visibility")
            End With
            colWorkbookWorksheets.Add objWorkbookWorksheet
        Next
    End If
    Set FetchWorksheets = colWorkbookWorksheets
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
            End With
            colWorkbookTables.Add objWorkbookTable
        Next
    End If
    Set FetchTables = colWorkbookTables
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
