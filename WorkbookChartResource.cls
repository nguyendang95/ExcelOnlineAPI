VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookChartResource"
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

Public Function ListCharts(ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As Collection
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
    strUrl = strUrl & "/charts"
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
            Set ListCharts = FetchWorkbookCharts(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Function FetchWorkbookCharts(ByVal ResponseText As String) As Collection
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
            colWorkbookCharts.Add objWorkbookChart
        Next
    End If
    Set FetchWorkbookCharts = colWorkbookCharts
End Function

Public Function GetChart(ByVal ChartName As String, ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As WorkbookChart
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
    strUrl = strUrl & "/charts/" & ChartName
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
            Set GetChart = FetchWorkbookChart(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function UpdateChart(ByVal ChartName As String, ByVal WorkbookChart As WorkbookChart, ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookChart
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
    strUrl = strUrl & "/charts/" & ChartName
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(WorkbookChart.ToJson)
        .WaitForResponse
        If .Status = 200 Then
            Set UpdateChart = FetchWorkbookChart(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function DeleteChart(ByVal ChartName As String, ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookChart
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
    strUrl = strUrl & "/charts/" & ChartName
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
End Function

Public Function AddChart(ByVal WorkbookChart As WorkbookChart, ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookChart
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
    strUrl = strUrl & "/charts/add"
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(WorkbookChart.ToJson)
        .WaitForResponse
        If .Status = 201 Then
            Set AddChart = FetchWorkbookChart(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetChartImage(ByVal ChartName As String, ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal PathParameters As String) As String
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
    strUrl = strUrl & "/charts/" & ChartName & "/image"
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            GetChartImage = objJson.Item("value")
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub ResetData(ByVal ChartName As String, ByVal SourceData As Variant, ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal SeriesBy As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
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
    strUrl = strUrl & "/charts/" & ChartName & "/setData"
    Dim objRequestBody As Scripting.Dictionary
    Set objRequestBody = New Scripting.Dictionary
    With objRequestBody
        .Add "sourceData", JsonConverter.ParseJson(JsonConverter.ConvertToJson(SourceData))
        If SeriesBy <> vbNullString Then
            .Add "seriesBy", SeriesBy
        Else: .Add "seriesBy", "Auto"
        End If
    End With
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(objRequestBody)
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Sub SetPosition(ByVal ChartName As String, ByVal StartCell As Variant, ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal EndCell As Variant, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
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
    strUrl = strUrl & "/charts/" & ChartName & "/setPosition"
    Dim objRequestBody As Scripting.Dictionary
    Set objRequestBody = New Scripting.Dictionary
    With objRequestBody
        .Add "startCell", JsonConverter.ParseJson(JsonConverter.ConvertToJson(StartCell))
        If IsMissing(EndCell) = False Then .Add "endCell", JsonConverter.ParseJson(JsonConverter.ConvertToJson(EndCell))
    End With
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(objRequestBody)
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Function GetChartSeries(ByVal WorkbookChartSeries As WorkbookChartSeries, ByVal ChartName As String, ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookChartSeries
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
    strUrl = strUrl & "/charts/" & ChartName & "/series"
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(WorkbookChartSeries.ToJson)
        .WaitForResponse
        If .Status = 201 Then
            Set GetChartSeries = FetchWorkbookChartSeries(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListChartSeriesCollection(ByVal ChartName As String, ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As Collection
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
    strUrl = strUrl & "/charts/" & ChartName & "/series"
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
            Set ListChartSeriesCollection = FetchWorkbookChartSeriesCollection(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Function FetchWorkbookChartSeries(ByVal ResponseText As String) As WorkbookChartSeries
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookChartSeries As WorkbookChartSeries
    Set objWorkbookChartSeries = New WorkbookChartSeries
    With objWorkbookChartSeries
        .Name = objJson.Item("name")
    End With
    Set FetchWorkbookChartSeries = objWorkbookChartSeries
End Function

Private Function FetchWorkbookChartSeriesCollection(ByVal ResponseText As String) As Collection
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookChartSeries As WorkbookChartSeries
    Dim colWorkbookChartSeries As New Collection
    Dim i As Long
    If objJson.Item("value").Count > 0 Then
        For i = 1 To objJson.Item("value").Count
            Set objWorkbookChartSeries = New WorkbookChartSeries
            With objWorkbookChartSeries
                .Name = objJson.Item("value")(i)("name")
            End With
            colWorkbookChartSeries.Add objWorkbookChartSeries
        Next
    End If
    Set FetchWorkbookChartSeriesCollection = colWorkbookChartSeries
End Function

Private Function FetchWorkbookChart(ByVal ResponseText As String) As WorkbookChart
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
    Set FetchWorkbookChart = objWorkbookChart
End Function

Private Sub HandleError(Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
