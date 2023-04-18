VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookApplicationResource"
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

Public Function GetWorkbookApplication(ByVal Location As Location, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As WorkbookApplication
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/application"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/functions/application"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/application"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/functions/application"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/application"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/functions/application"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/application"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/functions/application"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/application"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/functions/application"
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
            Set GetWorkbookApplication = FetchWorkbookApplication(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub Calculate(Location As Location, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Location
        Case LocationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/application"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/workbook/functions/application"
            End If
        Case LocationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/application"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/workbook/functions/application"
            End If
        Case LocationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/application"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/workbook/functions/application"
            End If
        Case LocationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/application"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/workbook/functions/application"
            End If
        Case LocationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/workbook/functions/application"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/workbook/functions/application"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        .WaitForResponse
        If .Status <> 200 Then Call HandleError(.ResponseText)
    End With
End Sub

Private Function FetchWorkbookApplication(ByVal ResponseText As String) As WorkbookApplication
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objWorkbookApplication As WorkbookApplication
    Set objWorkbookApplication = New WorkbookApplication
    objWorkbookApplication.CalculationMode = objJson.Item("calculationMode")
    Set FetchWorkbookApplication = objWorkbookApplication
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub
