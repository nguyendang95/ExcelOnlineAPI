VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookCommentResource"
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

Public Function GetWorkbookComment(Location As Location, ByVal CommentId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookComment
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/comments/" & CommentId
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/functions/comments/" & CommentId
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetWorkbookComment = FetchWorkbookComment(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListWorkbookComments(Location As Location, ByVal CommentId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/comments/" & CommentId
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/functions/comments/" & CommentId
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set ListWorkbookComments = FetchWorkbookComments(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListReplies(Location As Location, ByVal CommentId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/comments/" & CommentId & "/replies"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId & "/replies"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId & "/replies"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId & "/replies"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId & "/replies"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set ListReplies = FetchWorkbookCommentReplies(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetReply(Location As Location, ByVal CommentId As String, ReplyId As String, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/comments/" & CommentId & "/replies/" & ReplyId
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies/" & ReplyId
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId & "/replies/" & ReplyId
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies/" & ReplyId
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId & "/replies/" & ReplyId
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies/" & ReplyId
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId & "/replies/" & ReplyId
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies/" & ReplyId
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId & "/replies/" & ReplyId
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies/" & ReplyId
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetReply = FetchWorkbookCommentReply(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function CreateReply(Location As Location, ByVal CommentId As String, WorkbookCommentReply As WorkbookCommentReply, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/comments/" & CommentId & "/replies/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies/"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId & "/replies/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies/"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId & "/replies/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies/"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId & "/replies/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies/"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/comments/" & CommentId & "/replies/"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/functions/comments/" & CommentId & "/replies/"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(WorkbookCommentReply.ToJson)
        .WaitForResponse
        If .Status = 201 Then
            Set CreateReply = FetchWorkbookCommentReply(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Function FetchWorkbookComment(ByVal ResponseText As String) As WorkbookComment
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookComment As WorkbookComment
    Set objWorkbookComment = New WorkbookComment
    With objWorkbookComment
        .Content = objJson.Item("content")
        .ContentType = objJson.Item("contentType")
        .Id = objJson.Item("id")
    End With
    Set FetchWorkbookComment = objWorkbookComment
End Function

Private Function FetchWorkbookComments(ByVal ResponseText As String) As Collection
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookComment As WorkbookComment
    Dim i As Long
    Dim colWorkbookComments As New Collection
    If objJson.Item("value").Count > 0 Then
        For i = 1 To objJson.Item("value").Count
            Set objWorkbookComment = New WorkbookComment
            With objWorkbookComment
                .Content = objJson.Item("value")(i)("content")
                .ContentType = objJson.Item("value")(i)("contentType")
                .Id = objJson.Item("value")(i)("id")
            End With
            colWorkbookComments.Add objWorkbookComment
        Next
    End If
    Set FetchWorkbookComments = colWorkbookComments
End Function

Private Function FetchWorkbookCommentReply(ByVal ResponseText As String) As WorkbookCommentReply
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookCommentReply As WorkbookCommentReply
    Set objWorkbookCommentReply = New WorkbookCommentReply
    With objWorkbookCommentReply
        .Content = objJson.Item("content")
        .ContentType = objJson.Item("contentType")
        .Id = objJson.Item("id")
    End With
    Set FetchWorkbookCommentReply = objWorkbookCommentReply
End Function

Private Function FetchWorkbookCommentReplies(ByVal ResponseText As String) As Collection
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookCommentReply As WorkbookCommentReply
    Dim i As Long
    Dim colWorkbookCommentReplies As New Collection
    If objJson.Item("value").Count > 0 Then
        For i = 1 To objJson.Item("value").Count
            Set objWorkbookCommentReply = New WorkbookCommentReply
            With objWorkbookCommentReply
                .Content = objJson.Item("value")(i)("content")
                .ContentType = objJson.Item("value")(i)("contentType")
                .Id = objJson.Item("value")(i)("id")
            End With
            colWorkbookCommentReplies.Add objWorkbookCommentReply
        Next
    End If
    Set FetchWorkbookCommentReplies = colWorkbookCommentReplies
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
