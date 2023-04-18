VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MicrosoftGraphOAuth2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum TenantType
    Common
    Organizations
    Consumers
End Enum

Private m_ApplicationName As String
Private m_ClientId As String
Private m_ClientSecret As String
Private m_Scope As Variant
Private m_AccessToken As String
Private m_Tenant As TenantType
Private m_WorkbookResource As WorkbookResource
Private m_WorkbookApplicationResource As WorkbookApplicationResource
Private m_WorksheetResource As WorksheetResource
Private m_WorkbookCommentResource As WorkbookCommentResource
Private m_WorkbookTableRowResource As WorkbookTableRowResource
Private m_WorkbookRangeResource As WorkbookRangeResource
Private m_WorkbookPivotTableResource As WorkbookPivotTableResource
Private m_WorkbookTableResource As WorkbookTableResource
Private m_WorkbookTableColumnResource As WorkbookTableColumnResource
Private m_WorkbookFilterResource As WorkbookFilterResource
Private m_WorkbookChartResource As WorkbookChartResource
Private m_WorkbookChartSeriesResource As WorkbookChartSeriesResource
Private m_WorkbookChartPointsResource As WorkbookChartPointsResource
Private m_WorkbookNamedItemResource As WorkbookNamedItemResource

Public Property Get WorkbookNamedItemResource() As WorkbookNamedItemResource
   Set WorkbookNamedItemResource = m_WorkbookNamedItemResource
End Property

Public Property Set WorkbookNamedItemResource(Value As WorkbookNamedItemResource)
   Set m_WorkbookNamedItemResource = Value
End Property

Public Property Get WorkbookChartPointsResource() As WorkbookChartPointsResource
   Set WorkbookChartPointsResource = m_WorkbookChartPointsResource
End Property

Public Property Set WorkbookChartPointsResource(Value As WorkbookChartPointsResource)
   Set m_WorkbookChartPointsResource = Value
End Property

Public Property Get WorkbookChartSeriesResource() As WorkbookChartSeriesResource
   Set WorkbookChartSeriesResource = m_WorkbookChartSeriesResource
End Property

Public Property Set WorkbookChartSeriesResource(Value As WorkbookChartSeriesResource)
   Set m_WorkbookChartSeriesResource = Value
End Property

Public Property Get WorkbookChartResource() As WorkbookChartResource
   Set WorkbookChartResource = m_WorkbookChartResource
End Property

Public Property Set WorkbookChartResource(Value As WorkbookChartResource)
   Set m_WorkbookChartResource = Value
End Property

Public Property Get WorkbookFilterResource() As WorkbookFilterResource
   Set WorkbookFilterResource = m_WorkbookFilterResource
End Property

Public Property Set WorkbookFilterResource(Value As WorkbookFilterResource)
   Set m_WorkbookFilterResource = Value
End Property

Public Property Get WorkbookTableColumnResource() As WorkbookTableColumnResource
   Set WorkbookTableColumnResource = m_WorkbookTableColumnResource
End Property

Public Property Set WorkbookTableColumnResource(Value As WorkbookTableColumnResource)
   Set m_WorkbookTableColumnResource = Value
End Property

Public Property Get WorkbookTableResource() As WorkbookTableResource
   Set WorkbookTableResource = m_WorkbookTableResource
End Property

Public Property Set WorkbookTableResource(Value As WorkbookTableResource)
   Set m_WorkbookTableResource = Value
End Property

Public Property Get WorkbookPivotTableResource() As WorkbookPivotTableResource
   Set WorkbookPivotTableResource = m_WorkbookPivotTableResource
End Property

Public Property Set WorkbookPivotTableResource(Value As WorkbookPivotTableResource)
   Set m_WorkbookPivotTableResource = Value
End Property

Public Property Get WorkbookRangeResource() As WorkbookRangeResource
   Set WorkbookRangeResource = m_WorkbookRangeResource
End Property

Public Property Set WorkbookRangeResource(Value As WorkbookRangeResource)
   Set m_WorkbookRangeResource = Value
End Property

Public Property Get WorkbookTableRowResource() As WorkbookTableRowResource
   Set WorkbookTableRowResource = m_WorkbookTableRowResource
End Property

Public Property Set WorkbookTableRowResource(Value As WorkbookTableRowResource)
   Set m_WorkbookTableRowResource = Value
End Property

Public Property Get WorkbookCommentResource() As WorkbookCommentResource
   Set WorkbookCommentResource = m_WorkbookCommentResource
End Property

Public Property Set WorkbookCommentResource(Value As WorkbookCommentResource)
   Set m_WorkbookCommentResource = Value
End Property

Public Property Get WorksheetResource() As WorksheetResource
   Set WorksheetResource = m_WorksheetResource
End Property

Public Property Set WorksheetResource(Value As WorksheetResource)
   Set m_WorksheetResource = Value
End Property

Public Property Get WorkbookApplicationResource() As WorkbookApplicationResource
   Set WorkbookApplicationResource = m_WorkbookApplicationResource
End Property

Public Property Set WorkbookApplicationResource(Value As WorkbookApplicationResource)
   Set m_WorkbookApplicationResource = Value
End Property

Public Property Get WorkbookResource() As WorkbookResource
   Set WorkbookResource = m_WorkbookResource
End Property

Public Property Set WorkbookResource(Value As WorkbookResource)
   Set m_WorkbookResource = Value
End Property


Public Property Get ApplicationName() As String
    ApplicationName = m_ApplicationName
End Property

Public Property Let ApplicationName(Value As String)
    m_ApplicationName = Value
End Property

Public Property Get ClientId() As String
    ClientId = m_ClientId
End Property

Public Property Let ClientId(Value As String)
    m_ClientId = Value
End Property

Public Property Get ClientSecret() As String
    ClientSecret = m_ClientSecret
End Property

Public Property Let ClientSecret(Value As String)
    m_ClientSecret = Value
End Property

Public Property Get Scope() As Variant
    Scope = m_Scope
End Property

Public Property Let Scope(Value As Variant)
    m_Scope = Value
End Property

Public Property Get AccessToken() As String
    AccessToken = m_AccessToken
End Property

Public Property Let AccessToken(Value As String)
    m_AccessToken = Value
End Property

Public Property Get Tenant() As TenantType
    Tenant = m_Tenant
End Property

Public Property Let Tenant(Value As TenantType)
    m_Tenant = Value
End Property

Public Sub AuthorizeOAuth2()
    Dim objWinHttp As WinHttp.WinHttpRequest
    Dim strRequestBody As String
    Dim strUrl As String
    Dim objReg As RegistryUtility
    Dim strAuthorizationCode As String
    Dim objFSO As Scripting.FileSystemObject
    Dim objFolder As Scripting.Folder
    Dim objFile As Scripting.TextStream
    Dim strFolder As String
    Dim strFile As String
    Dim strAccessToken As String
    Dim objJson As Scripting.Dictionary
    Dim dteExpTime As Date
    Dim strRefreshToken As String
    strFolder = Environ$("APPDATA") & "\" & ApplicationName
    strFile = strFolder & "\client_id_" & ClientId & ".token-response"
    Set objReg = New RegistryUtility
    Set objFSO = New Scripting.FileSystemObject
    If ApplicationName = vbNullString Then
        Err.Raise vbObjectError + 10, , "Please specify a name for your application"
    End If
    If objFSO.FolderExists(strFolder) Then
        Set objFolder = objFSO.GetFolder(strFolder)
        If objFSO.FileExists(strFile) Then
            If objReg.RegValueExists("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AccessTokenExpirationTime") Then
                dteExpTime = CDate(objReg.ReadRegValue("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AccessTokenExpirationTime"))
                If DateDiff("s", Now, dteExpTime) <= 0 Then
                    Set objFile = objFSO.OpenTextFile(strFile, ForReading, False)
                    Set objJson = JsonConverter.ParseJson(objFile.ReadAll)
                    objFile.Close
                    strRefreshToken = objJson.Item("refresh_token")
                    AccessToken = RefreshAccessToken(strRefreshToken)
                Else
                    Set objFile = objFSO.OpenTextFile(strFile, ForReading, False)
                    Set objJson = JsonConverter.ParseJson(objFile.ReadAll)
                    objFile.Close
                    AccessToken = objJson.Item("access_token")
                End If
            Else: AccessToken = GetAccessToken
            End If
        Else: AccessToken = GetAccessToken
        End If
    Else
        Set objFolder = objFSO.CreateFolder(strFolder)
        AccessToken = GetAccessToken
    End If
    Dim objWorkbookResource As WorkbookResource
    Set objWorkbookResource = New WorkbookResource
    objWorkbookResource.AccessToken = AccessToken
    Set WorkbookResource = objWorkbookResource
    Dim objWorkbookApplicationResource As WorkbookApplicationResource
    Set objWorkbookApplicationResource = New WorkbookApplicationResource
    objWorkbookApplicationResource.AccessToken = AccessToken
    Set WorkbookApplicationResource = objWorkbookApplicationResource
    Dim objWorksheetResource As WorksheetResource
    Set objWorksheetResource = New WorksheetResource
    objWorksheetResource.AccessToken = AccessToken
    Set WorksheetResource = objWorksheetResource
    Dim objWorkbookCommentResource As WorkbookCommentResource
    Set objWorkbookCommentResource = New WorkbookCommentResource
    objWorkbookCommentResource.AccessToken = AccessToken
    Set WorkbookCommentResource = objWorkbookCommentResource
    Dim objWorkbookTableRowResource As WorkbookTableRowResource
    Set objWorkbookTableRowResource = New WorkbookTableRowResource
    objWorkbookTableRowResource.AccessToken = AccessToken
    Set WorkbookTableRowResource = objWorkbookTableRowResource
    Dim objWorkbookRangeResource As WorkbookRangeResource
    Set objWorkbookRangeResource = New WorkbookRangeResource
    objWorkbookRangeResource.AccessToken = AccessToken
    Set WorkbookRangeResource = objWorkbookRangeResource
    Dim objWorkbookPivotTableResource As WorkbookPivotTableResource
    Set objWorkbookPivotTableResource = New WorkbookPivotTableResource
    objWorkbookPivotTableResource.AccessToken = AccessToken
    Set WorkbookPivotTableResource = objWorkbookPivotTableResource
    Dim objWorkbookTableResource As WorkbookTableResource
    Set objWorkbookTableResource = New WorkbookTableResource
    objWorkbookTableResource.AccessToken = AccessToken
    Set WorkbookTableResource = objWorkbookTableResource
    Dim objWorkbookTableColumnResource As WorkbookTableColumnResource
    Set objWorkbookTableColumnResource = New WorkbookTableColumnResource
    objWorkbookTableColumnResource.AccessToken = AccessToken
    Set WorkbookTableColumnResource = objWorkbookTableColumnResource
    Dim objWorkbookFilterResource As WorkbookFilterResource
    Set objWorkbookFilterResource = New WorkbookFilterResource
    objWorkbookFilterResource.AccessToken = AccessToken
    Set WorkbookFilterResource = objWorkbookFilterResource
    Dim objWorkbookChartResource As WorkbookChartResource
    Set objWorkbookChartResource = New WorkbookChartResource
    objWorkbookChartResource.AccessToken = AccessToken
    Set WorkbookChartResource = objWorkbookChartResource
    Dim objWorkbookChartSeriesResource As WorkbookChartSeriesResource
    Set objWorkbookChartSeriesResource = New WorkbookChartSeriesResource
    objWorkbookChartSeriesResource.AccessToken = AccessToken
    Set WorkbookChartSeriesResource = objWorkbookChartSeriesResource
    Dim objWorkbookChartPointsResource As WorkbookChartPointsResource
    Set objWorkbookChartPointsResource = New WorkbookChartPointsResource
    objWorkbookChartPointsResource.AccessToken = AccessToken
    Set WorkbookChartPointsResource = objWorkbookChartPointsResource
    Dim objWorkbookNamedItemResource As WorkbookNamedItemResource
    Set objWorkbookNamedItemResource = New WorkbookNamedItemResource
    objWorkbookNamedItemResource.AccessToken = AccessToken
    Set WorkbookNamedItemResource = objWorkbookNamedItemResource
End Sub

Private Function GetAccessToken() As String
    Dim objBrowser As UserForm1
    Dim objWinHttp As WinHttp.WinHttpRequest
    Dim strUrl As String
    Dim strRequestBody As String
    Dim strAuthorizationCode As String
    Dim strFile As String
    Dim strFolder As String
    Dim strAccessToken As String
    Dim objReg As RegistryUtility
    Dim objJson As Scripting.Dictionary
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.TextStream
    Set objFSO = New Scripting.FileSystemObject
    Set objReg = New RegistryUtility
    Set objBrowser = New UserForm1
    strFolder = Environ$("APPDATA") & "\" & ApplicationName
    strFile = strFolder & "\client_id_" & ClientId & ".token-response"
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strState As String
    strState = GenerateRandomString(128)
    Dim strTenant As String
    Select Case Tenant
        Case Common
            strTenant = "common"
        Case Organizations
            strTenant = "organizations"
        Case Consumers
            strTenant = "consumers"
        Case Else: Err.Raise vbObjectError + 14, , "Wrong value or empty tenant value!"
    End Select
    If Not IsEmpty(Scope) Then
        strUrl = "https://login.microsoftonline.com/" & strTenant & "/oauth2/v2.0/authorize?" & "client_id=" & ClientId & "&scope=" & objWebUtilities.URLEncode("offline_access ") & objWebUtilities.URLEncode(Join(Scope, " ")) & "&response_type=code&redirect_uri=" & objWebUtilities.URLEncode("https://login.microsoftonline.com/common/oauth2/nativeclient") & "&state=" & strState & "&response_mode=query"
    Else
        Err.Raise vbObjectError + 13, , "Scope cannot be null. Please specify a valid scope and try again!"
    End If
    objBrowser.ClientId = ClientId
    objBrowser.ApplicationName = ApplicationName
    objBrowser.Mode = "Login"
    objBrowser.WebBrowser1.Navigate strUrl
    objBrowser.Show vbModal
    If objReg.RegValueExists("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\ServerSideStateCode") Then
        If strState <> objReg.ReadRegValue("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\ServerSideStateCode") Then
            objReg.DeleteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\ServerSideStateCode"
            objReg.DeleteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AuthorizationCode"
            Err.Raise vbObjectError, , "Malformed authorization response!"
        End If
        objReg.DeleteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\ServerSideStateCode"
    End If
    If objReg.RegValueExists("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AuthorizationCode") Then
        strAuthorizationCode = objReg.ReadRegValue("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AuthorizationCode")
        objReg.DeleteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AuthorizationCode"
    End If
    If strAuthorizationCode <> vbNullString Then
        strRequestBody = "code=" & strAuthorizationCode & "&" & _
                        "state=" & strState & "&" & _
                        "client_id=" & ClientId & "&" & _
                        "redirect_uri=" & objWebUtilities.URLEncode("https://login.microsoftonline.com/common/oauth2/nativeclient") & "&" & _
                        "grant_type=authorization_code"
        Set objWinHttp = New WinHttp.WinHttpRequest
        With objWinHttp
            .Open "POST", "https://login.microsoftonline.com/" & strTenant & "/oauth2/v2.0/token", False
            .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .SetRequestHeader "Accept", "application/json'"
            .Send strRequestBody
            If .Status = 200 Then
                Set objJson = JsonConverter.ParseJson(.ResponseText)
                strAccessToken = objJson.Item("access_token")
                If strAccessToken <> vbNullString Then
                    GetAccessToken = strAccessToken
                    Set objFile = objFSO.CreateTextFile(strFile, True)
                    objFile.Write .ResponseText
                    objFile.Close
                    objReg.WriteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AccessTokenExpirationTime", CStr(DateAdd("s", CDbl(objJson.Item("expires_in")), Now)), REG_SZ
                    objReg.WriteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\Tenant", LCase(strTenant), REG_SZ
                Else: Err.Raise vbObjectError + 2, , "Failed to get access code"
                End If
            End If
        End With
    Else
        Err.Raise vbObjectError + 1, , "Failed to obtain the authorization code."
    End If
End Function

Private Function RefreshAccessToken(ByVal RefreshToken As String) As String
    Dim objWinHttp As WinHttp.WinHttpRequest
    Dim objJson As Scripting.Dictionary
    Dim strRequestBody As String
    Dim strFolder As String
    Dim strFile As String
    Dim objReg As RegistryUtility
    Set objReg = New RegistryUtility
    Dim strTenant As String
    strTenant = objReg.ReadRegValue("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\Tenant")
    strRequestBody = "client_id=" & ClientId & "&" & _
                    "refresh_token=" & RefreshToken & "&" & _
                    "grant_type=refresh_token"
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", "https://login.microsoftonline.com/" & strTenant & "/oauth2/v2.0/token"
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .SetRequestHeader "Accept", "application/json'"
        .Send strRequestBody
        If .Status = 200 Then
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            strFolder = Environ$("APPDATA") & "\" & ApplicationName
            strFile = strFolder & "\client_id_" & ClientId & ".token-response"
            Call SaveNewAccessToken(strFile, objJson.Item("access_token"))
            RefreshAccessToken = objJson.Item("access_token")
            objReg.WriteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AccessTokenExpirationTime", CStr(DateAdd("s", CDbl(objJson.Item("expires_in")), Now)), REG_SZ
        End If
    End With
End Function

Private Sub SaveNewAccessToken(ByVal AccessTokenFile As String, ByVal NewAccessToken As String)
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.TextStream
    Dim strText As String, strOldText As String, strNewText As String
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.FileExists(AccessTokenFile) Then
        strNewText = Chr(34) & "access_token" & Chr(34) & ":" & Chr(34) & NewAccessToken & Chr(34)
        Set objFile = objFSO.OpenTextFile(AccessTokenFile)
        strText = objFile.ReadAll
        objFile.Close
        strOldText = Chr(34) & "access_token" & Chr(34) & ":" & Chr(34) & JsonConverter.ParseJson(strText).Item("access_token") & Chr(34)
        strText = VBA.Replace(strText, strOldText, strNewText, , , vbTextCompare)
        Set objFile = objFSO.OpenTextFile(AccessTokenFile, ForWriting)
        objFile.Write strText
        objFile.Close
    Else
        Call GetAccessToken
    End If
End Sub

Public Sub LogOut()
    Dim objWinHttp As WinHttp.WinHttpRequest
    Dim objFSO As Scripting.FileSystemObject
    Dim strFolder As String
    Dim objReg As RegistryUtility
    Dim objBrowser As UserForm1
    Set objBrowser = New UserForm1
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strTenant As String
    Set objReg = New RegistryUtility
    If objReg.RegValueExists("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\Tenant") Then
        strTenant = objReg.ReadRegValue("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\Tenant")
    Else: Exit Sub
    End If
    If strTenant <> vbNullString Then objBrowser.WebBrowser1.Navigate "https://login.microsoftonline.com/" & strTenant & "/oauth2/v2.0/logout?post_logout_redirect_uri=" & objWebUtilities.URLEncode("https://login.microsoftonline.com/common/oauth2/nativeclient")
    objBrowser.Show vbModal
    Set objFSO = New Scripting.FileSystemObject
    strFolder = Environ$("APPDATA") & "\" & ApplicationName
    If objFSO.FileExists(strFolder & "\client_id_" & ClientId & ".token-response") Then objFSO.GetFile(strFolder & "\client_id_" & ClientId & ".token-response").Delete
    If objReg.RegKeyExists(HKEY_CURRENT_USER, "Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId) Then
        objReg.DeleteRegKey HKEY_CURRENT_USER, "Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId
    End If
End Sub

Private Function GenerateRandomString(ByVal Length As Long) As String
    Dim i As Long
    Dim strText As String
    If Length <= 0 Then Err.Raise vbObjectError, , "Length of string cannot below zero."
    Dim arrSpecialCharacters As Variant
    arrSpecialCharacters = Array("-", ".", "_", "~")
    Randomize
    For i = 1 To Length
        If RandBetween(0, 3) = 0 Then
            strText = strText & Chr(RandBetween(65, 90))
        ElseIf RandBetween(0, 3) = 1 Then strText = strText & Chr(RandBetween(48, 57))
        ElseIf RandBetween(0, 3) = 2 Then strText = strText & arrSpecialCharacters(RandBetween(0, 3))
        Else: strText = strText & Chr(RandBetween(97, 122))
        End If
    Next
    GenerateRandomString = strText
End Function

Private Function RandBetween(ByVal Num1 As Long, ByVal Num2 As Long) As Long
    RandBetween = CLng((Num2 - Num1) * Rnd + Num1)
End Function
