VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookTableResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum TableLocation
    InWorkbook
    InWorksheet
    InTables
End Enum

Private m_AccessToken As String

Friend Property Get AccessToken() As String
   AccessToken = m_AccessToken
End Property

Friend Property Let AccessToken(Value As String)
   m_AccessToken = Value
End Property

Public Function Add(ByVal Address As String, ByVal HasHeader As Boolean, TableLocation As TableLocation, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As WorkbookTable
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/add"
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/add"
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/add"
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/add"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/add"
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/add"
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/add"
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/add"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/add"
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/add"
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/add"
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/add"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/add"
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/add"
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/add"
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/add"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/add"
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/tables/add"
                End Select
            Else
                Select Case TableLocation
                    Case InWorkbook
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/add"
                    Case InWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/tables/add"
                End Select
            End If
    End Select
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objRequestBody As Scripting.Dictionary
    Set objRequestBody = New Scripting.Dictionary
    objRequestBody.Add "address", Address
    objRequestBody.Add "hasHeader", HasHeader
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
            Set Add = FetchTable(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetTable(ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookTable
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
            Set GetTable = FetchTable(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function Update(ByVal WorkbookTable As WorkbookTable, ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookTable
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
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "PATCH", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(WorkbookTable)
        .WaitForResponse
        If .Status = 200 Then
            Set Update = FetchTable(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub Delete(ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
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

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
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

Public Function ListRows(ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As Collection
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
    strUrl = strUrl & "/rows"
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
            Set ListRows = FetchWorkbookTableRows(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function CreateRow(ByVal Values As Variant, ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal Index As Variant, Optional ByVal UseAsyncRequest As Boolean = False, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As Variant
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
    strUrl = strUrl & "/rows"
    Dim objRequestBody As Scripting.Dictionary
    Set objRequestBody = New Scripting.Dictionary
    With objRequestBody
        .Add "values", JsonConverter.ParseJson(JsonConverter.ConvertToJson(Values))
        If IsMissing(Index) = False Then
            .Add "index", Index
        End If
    End With
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If UseAsyncRequest Then .SetRequestHeader "Prefer", "respond-async"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(objRequestBody)
        .WaitForResponse
        If .Status = 202 Then CreateRow = .GetResponseHeader("Location")
        If .Status = 201 Then
            Set CreateRow = FetchWorkbookTableRow(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub DeleteRow(ByVal RowIndex As Long, ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
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
    strUrl = strUrl & "/rows/itemAt(index=" & RowIndex & ")"
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

Public Function ListColumns(ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As Collection
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

Public Sub DeleteColumn(ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
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
        .Open "DELETE", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Function GetRange(ByVal TableIdOrName As String, TableLocation As TableLocation, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
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
    strUrl = strUrl & "/range"
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
    Dim objWorkbookTableRow As WorkbookTableRow
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
