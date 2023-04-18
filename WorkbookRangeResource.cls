VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookRangeResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum RangeType
    RangeTypeNamedItem
    RangeTypeWorksheet
    RangeTypeTableColumn
End Enum

Private m_AccessToken As String

Friend Property Get AccessToken() As String
   AccessToken = m_AccessToken
End Property

Friend Property Let AccessToken(Value As String)
   m_AccessToken = Value
End Property

Public Function GetRange(RangeType As RangeType, Location As Location, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
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
            Set GetRange = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub DeleteRange(RangeType As RangeType, Location As Location, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal Address As String, Optional ByVal Shift As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/delete"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/delete"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/delete"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/delete"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/delete"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/delete"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/delete"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/delete"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/delete"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/delete"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/delete"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/delete"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/delete"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/delete"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/delete"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/delete"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/delete"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/delete"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/delete"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/delete"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/delete"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/delete"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/delete"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/delete"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/delete"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/delete"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/delete"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/delete"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/delete"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/delete"
                End Select
            End If
    End Select
    If Shift <> vbNullString Then
        Dim objRequestBody As Scripting.Dictionary
        Set objRequestBody = New Scripting.Dictionary
        objRequestBody.Add "shift", Shift
    End If
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        If Shift <> vbNullString Then
            .Send JsonConverter.ConvertToJson(objRequestBody)
        Else: .Send
        End If
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Sub ClearRange(RangeType As RangeType, Location As Location, Optional ByVal ApplyTo As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal Address As String, Optional ByVal Shift As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/clear"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/clear"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/clear"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/clear"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/clear"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/clear"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/clear"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/clear"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/clear"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/clear"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/clear"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/clear"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/clear"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/clear"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/clear"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/clear"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/clear"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/clear"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/clear"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/clear"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/clear"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/clear"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/clear"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/clear"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/clear"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/clear"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/clear"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/clear"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/clear"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/clear"
                End Select
            End If
    End Select
     If ApplyTo <> vbNullString Then
        Dim objRequestBody As Scripting.Dictionary
        Set objRequestBody = New Scripting.Dictionary
        objRequestBody.Add "applyTo", ApplyTo
    End If
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        If ApplyTo <> vbNullString Then
            .Send JsonConverter.ConvertToJson(objRequestBody)
        Else: .Send
        End If
        .WaitForResponse
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Sub UnmergeRange(RangeType As RangeType, Location As Location, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal Address As String, Optional ByVal Shift As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/unmerge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/unmerge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/unmerge"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/unmerge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/unmerge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/unmerge"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/unmerge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/unmerge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/unmerge"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/unmerge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/unmerge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/unmerge"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/unmerge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/unmerge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/unmerge"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/unmerge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/unmerge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/unmerge"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/unmerge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/unmerge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/unmerge"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/unmerge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/unmerge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/unmerge"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/unmerge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/unmerge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/unmerge"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/unmerge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/unmerge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/unmerge"
                End Select
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
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Sub MergeRange(RangeType As RangeType, Location As Location, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal Address As String, Optional ByVal Shift As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/merge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/merge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/merge"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/merge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/merge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/merge"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/merge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/merge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/merge"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/merge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/merge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/merge"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/merge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/merge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/merge"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/merge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/merge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/merge"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/merge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/merge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/merge"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/merge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/merge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/merge"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/merge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/merge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/merge"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/merge"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/merge"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/merge"
                End Select
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
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Function UpdateRangeFormat(ByVal WorkbookRangeFormat As WorkbookRangeFormat, RangeType As RangeType, Location As Location, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRangeFormat
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
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
        .Send JsonConverter.ConvertToJson(WorkbookRangeFormat.ToJson)
        .WaitForResponse
        If .Status = 200 Then
            Set UpdateRangeFormat = FetchRangeFormat(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetRangeFormat(RangeType As RangeType, Location As Location, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As WorkbookRangeFormat
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/format"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/format"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/format"
                End Select
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
            Set GetRangeFormat = FetchRangeFormat(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function InsertRange(RangeType As RangeType, Location As Location, Optional ByVal Shift As String, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/insert"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/insert"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/insert"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/insert"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/insert"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/insert"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/insert"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/insert"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/insert"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/insert"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/insert"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/insert"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/insert"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/insert"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/insert"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/insert"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/insert"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/insert"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/insert"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/insert"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/insert"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/insert"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/insert"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/insert"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/insert"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/insert"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/insert"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/insert"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/insert"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/insert"
                End Select
            End If
    End Select
    If Shift <> vbNullString Then
    Dim objRequestBody As Scripting.Dictionary
        Set objRequestBody = New Scripting.Dictionary
        objRequestBody.Add "shift", Shift
    End If
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        If Shift <> vbNullString Then
            .Send JsonConverter.ConvertToJson(objRequestBody)
        Else: .Send
        End If
        .WaitForResponse
        If .Status = 200 Then
            Set InsertRange = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function UpdateRange(ByVal WorkbookRange As WorkbookRange, RangeType As RangeType, Location As Location, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range"
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
        .Send JsonConverter.ConvertToJson(WorkbookRange.ToJson)
        .WaitForResponse
        If .Status = 200 Then
            Set UpdateRange = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetBoundingRect(RangeType As RangeType, Location As Location, ByVal AnotherRange As String, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/boundingRect"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/boundingRect"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/boundingRect"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/boundingRect"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/boundingRect"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/boundingRect"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/boundingRect"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/boundingRect"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/boundingRect"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/boundingRect"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/boundingRect"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/boundingRect"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/boundingRect"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/boundingRect"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/boundingRect"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/boundingRect"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/boundingRect"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/boundingRect"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/boundingRect"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/boundingRect"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/boundingRect"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/boundingRect"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/boundingRect"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/boundingRect"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/boundingRect"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/boundingRect"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/boundingRect"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/boundingRect"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/boundingRect"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/boundingRect"
                End Select
            End If
    End Select
    Dim objRequestBody As Scripting.Dictionary
    Set objRequestBody = New Scripting.Dictionary
    objRequestBody.Add "anotherRange", AnotherRange
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(objRequestBody)
        .WaitForResponse
        If .Status = 200 Then
            Set GetBoundingRect = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetCell(RangeType As RangeType, ByVal Row As Long, ByVal Column As Long, Location As Location, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/cell(row=" & Row & ",column=" & Column & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/cell(row=" & Row & ",column=" & Column & ")"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/cell(row=" & Row & ",column=" & Column & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/cell(row=" & Row & ",column=" & Column & ")"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/cell(row=" & Row & ",column=" & Column & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/cell(row=" & Row & ",column=" & Column & ")"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/cell(row=" & Row & ",column=" & Column & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/cell(row=" & Row & ",column=" & Column & ")"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/cell(row=" & Row & ",column=" & Column & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/cell(row=" & Row & ",column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/cell(row=" & Row & ",column=" & Column & ")"
                End Select
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetCell = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetColumn(RangeType As RangeType, ByVal Column As Long, Location As Location, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/column(column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/column(column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/column(column=" & Column & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/column(column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/column(column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/column(column=" & Column & ")"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/column(column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/column(column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/column(column=" & Column & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/column(column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/column(column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/column(column=" & Column & ")"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/column(column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/column(column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/column(column=" & Column & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/column(column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/column(column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/column(column=" & Column & ")"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/column(column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/column(column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/column(column=" & Column & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/column(column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/column(column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/column(column=" & Column & ")"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/column(column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/column(column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/column(column=" & Column & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/column(column=" & Column & ")"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/column(column=" & Column & ")"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/column(column=" & Column & ")"
                End Select
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetColumn = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetColumnsAfter(ByVal Count As Long, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRangeView
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim RangeType As RangeType
    RangeType = RangeTypeWorksheet
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsAfter(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsAfter(count=" & Count & ")"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsAfter(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsAfter(count=" & Count & ")"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsAfter(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsAfter(count=" & Count & ")"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsAfter(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsAfter(count=" & Count & ")"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsAfter(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsAfter(count=" & Count & ")"
                End Select
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetColumnsAfter = FetchWorkbookRangeView(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetcolumnsBefore(ByVal Count As Long, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRangeView
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim RangeType As RangeType
    RangeType = RangeTypeWorksheet
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsBefore(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsBefore(count=" & Count & ")"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsBefore(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsBefore(count=" & Count & ")"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsBefore(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsBefore(count=" & Count & ")"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsBefore(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsBefore(count=" & Count & ")"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsBefore(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range/columnsBefore(count=" & Count & ")"
                End Select
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetcolumnsBefore = FetchWorkbookRangeView(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetEntireColumn(RangeType As RangeType, Location As Location, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/entireColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireColumn"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/entireColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireColumn"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/entireColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireColumn"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/entireColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireColumn"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/entireColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireColumn"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/entireColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireColumn"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/entireColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireColumn"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/entireColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireColumn"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/entireColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireColumn"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/entireColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireColumn"
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
            Set GetEntireColumn = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetEntireRow(RangeType As RangeType, Location As Location, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/entireRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireRow"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/entireRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireRow"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/entireRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireRow"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/entireRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireRow"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/entireRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireRow"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/entireRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireRow"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/entireRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireRow"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/entireRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireRow"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/entireRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireRow"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/entireRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/entireRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/entireRow"
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
            Set GetEntireRow = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetIntersection(RangeType As RangeType, Location As Location, ByVal AnotherRange As String, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/intersection"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/intersection"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/intersection"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/intersection"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/intersection"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/intersection"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/intersection"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/intersection"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/intersection"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/intersection"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/intersection"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/intersection"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/intersection"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/intersection"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/intersection"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/intersection"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/intersection"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/intersection"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/intersection"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/intersection"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/intersection"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/intersection"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/intersection"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/intersection"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/intersection"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/intersection"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/intersection"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/intersection"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/intersection"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/intersection"
                End Select
            End If
    End Select
    Dim objRequestBody As Scripting.Dictionary
    Set objRequestBody = New Scripting.Dictionary
    objRequestBody.Add "anotherRange", AnotherRange
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(objRequestBody)
        .WaitForResponse
        If .Status = 200 Then
            Set GetIntersection = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetlastCell(RangeType As RangeType, Location As Location, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastCell"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastCell"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastCell"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/lastCell"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastCell"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastCell"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastCell"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastCell"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastCell"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/lastCell"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastCell"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastCell"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastCell"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastCell"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastCell"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/lastCell"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastCell"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastCell"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastCell"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastCell"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastCell"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/lastCell"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastCell"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastCell"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastCell"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastCell"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastCell"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/lastCell"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastCell"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastCell"
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
            Set GetlastCell = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetLastColumn(RangeType As RangeType, Location As Location, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastColumn"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/lastColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastColumn"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastColumn"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/lastColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastColumn"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastColumn"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/lastColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastColumn"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastColumn"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/lastColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastColumn"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastColumn"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/lastColumn"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastColumn"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastColumn"
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
            Set GetLastColumn = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetLastRow(RangeType As RangeType, Location As Location, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastRow"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/lastRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastRow"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastRow"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/lastRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastRow"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastRow"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/lastRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastRow"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastRow"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/lastRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastRow"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/lastRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastRow"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/lastRow"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/lastRow"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/lastRow"
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
            Set GetLastRow = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetOffsetRange(ByVal RowOffset As Long, ByVal ColumnOffset As Long, RangeType As RangeType, Location As Location, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/offsetRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/offsetRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/offsetRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/offsetRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/offsetRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/offsetRange"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/offsetRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/offsetRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/offsetRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/offsetRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/offsetRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/offsetRange"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/offsetRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/offsetRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/offsetRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/offsetRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/offsetRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/offsetRange"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/offsetRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/offsetRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/offsetRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/offsetRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/offsetRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/offsetRange"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/offsetRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/offsetRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/offsetRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/offsetRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/offsetRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/offsetRange"
                End Select
            End If
    End Select
    Dim objRequestBody As Scripting.Dictionary
    Set objRequestBody = New Scripting.Dictionary
    With objRequestBody
        .Add "rowOffset", RowOffset
        .Add "columnOffset", ColumnOffset
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
        If .Status = 200 Then
            Set GetOffsetRange = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetRow(ByVal Row As Long, RangeType As RangeType, Location As Location, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/row"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/row"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/row"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/row"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/row"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/row"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/row"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/row"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/row"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/row"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/row"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/row"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/row"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/row"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/row"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/row"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/row"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/row"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/row"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/row"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/row"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/row"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/row"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/row"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/row"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/row"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/row"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/row"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/row"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/row"
                End Select
            End If
    End Select
    Dim objRequestBody As Scripting.Dictionary
    Set objRequestBody = New Scripting.Dictionary
    With objRequestBody
        .Add "row", Row
    End With
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(objRequestBody)
        .WaitForResponse
        If .Status = 200 Then
            Set GetRow = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetRowsAbove(ByVal Count As Long, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRangeView
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim RangeType As RangeType
    RangeType = RangeTypeWorksheet
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsAbove(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsAbove(count=" & Count & ")"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsAbove(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsAbove(count=" & Count & ")"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsAbove(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsAbove(count=" & Count & ")"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsAbove(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsAbove(count=" & Count & ")"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsAbove(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsAbove(count=" & Count & ")"
                End Select
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetRowsAbove = FetchWorkbookRangeView(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetRowsBelow(ByVal Count As Long, Location As Location, Optional ByVal WorksheetIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRangeView
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim RangeType As RangeType
    RangeType = RangeTypeWorksheet
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsBelow(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsBelow(count=" & Count & ")"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsBelow(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsBelow(count=" & Count & ")"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsBelow(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsBelow(count=" & Count & ")"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsBelow(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsBelow(count=" & Count & ")"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsBelow(count=" & Count & ")"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range/rowsBelow(count=" & Count & ")"
                End Select
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetRowsBelow = FetchWorkbookRangeView(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetUsedRange(RangeType As RangeType, Location As Location, Optional ByVal ValuesOnly As Boolean = False, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/usedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/usedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/usedRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/usedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/usedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/usedRange"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/usedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/usedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/usedRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/usedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/usedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/usedRange"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/usedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/usedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/usedRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/usedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/usedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/usedRange"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/usedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/usedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/usedRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/usedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/usedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/usedRange"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/usedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/usedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/usedRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/usedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/usedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/usedRange"
                End Select
            End If
    End Select
    If ValuesOnly Then strUrl = strUrl & "(valuesOnly=true)"
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
            Set GetUsedRange = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetResizedRange(ByVal DeltaRows As Long, ByVal DeltaColumns As Long, RangeType As RangeType, Location As Location, Optional ByVal Address As String, Optional ByVal NamedItemId As String, Optional ByVal WorksheetIdOrName As String, Optional ByVal TableIdOrName As String, Optional ByVal ColumnIdOrName As String, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRange
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/resizedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/resizedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/resizedRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/resizedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/resizedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/resizedRange"
                End Select
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/resizedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/resizedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/resizedRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/resizedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/resizedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/resizedRange"
                End Select
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/resizedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/resizedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/resizedRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/names/" & NamedItemId & "/range/resizedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/resizedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/resizedRange"
                End Select
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/resizedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/resizedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/resizedRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/resizedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/resizedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/resizedRange"
                End Select
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/names/" & NamedItemId & "/range/resizedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/resizedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/resizedRange"
                End Select
            Else
                Select Case RangeType
                    Case RangeTypeNamedItem
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/names/" & NamedItemId & "/range/resizedRange"
                    Case RangeTypeWorksheet
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/worksheets/" & WorksheetIdOrName & "/range(address='" & Address & "')/resizedRange"
                    Case RangeTypeTableColumn
                        strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/tables/" & "/workbook/tables/" & TableIdOrName & "/columns/" & ColumnIdOrName & "/range/resizedRange"
                End Select
            End If
    End Select
    strUrl = strUrl & "(deltaRows=" & DeltaRows & ",deltaColumns=" & DeltaColumns & ")"
    Dim objRequestBody As Scripting.Dictionary
    Set objRequestBody = New Scripting.Dictionary
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", AccessToken
        .SetRequestHeader "Accept", "application/json"
        If WorkbookSessionId <> vbNullString Then .SetRequestHeader "workbook-session-id", WorkbookSessionId
        .Send JsonConverter.ConvertToJson(objRequestBody)
        .WaitForResponse
        If .Status = 200 Then
            Set GetResizedRange = FetchRange(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetVisibleView(ByVal WorksheetNameOrId As String, Location As Location, Optional ByVal WorkbookSessionId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookRangeView
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
    strUrl = strUrl & "/visibleView"
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
            Set GetVisibleView = FetchWorkbookRangeView(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Function FetchWorkbookRangeView(ByVal ResponseText As String) As WorkbookRangeView
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookRangeView As WorkbookRangeView
    Set objWorkbookRangeView = New WorkbookRangeView
    With objWorkbookRangeView
        Set .CellAddresses = objJson.Item("cellAddresses")
        .Index = objJson.Item("index")
        .ColumnCount = objJson.Item("columnCount")
        .RowCount = objJson.Item("rowCount")
        Set .ValueTypes = objJson.Item("valueTypes")
        Set .Formulas = objJson.Item("formulas")
        Set .FormulasLocal = objJson.Item("formulasLocal")
        Set .FormulasR1C1 = objJson.Item("formulasR1C1")
        Set .Text = objJson.Item("text")
        Set .Values = objJson.Item("values")
    End With
    Set FetchWorkbookRangeView = objWorkbookRangeView
End Function

Private Function FetchRangeFormat(ByVal ResponseText As String) As WorkbookRangeFormat
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookRangeFormat As WorkbookRangeFormat
    Set objWorkbookRangeFormat = New WorkbookRangeFormat
    With objWorkbookRangeFormat
        .ColumnWidth = objJson.Item("columnWidth")
        .HorizontalAlignment = objJson.Item("horizontalAlignment")
        .RowHeight = objJson.Item("rowHeight")
        .VerticalAlignment = objJson.Item("verticalAlignment")
        .WrapText = objJson.Item("wrapText")
    End With
    Set FetchRangeFormat = objWorkbookRangeFormat
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
