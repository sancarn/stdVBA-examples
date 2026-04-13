VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} stdSharepointAuthenticator
   Caption         =   "Sharepoint Authenticator"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   OleObjectBlob   =   "stdSharepointAuthenticator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "stdSharepointAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'@description A class used to authenticate to a SharePoint site.
'@example `set auth = stdSharepointAuthenticator.Create("https://contoso.sharepoint.com")`

Option Explicit

Implements stdICallable

Private Const DEFAULT_READY_TIMEOUT_MS As Long = 180000
Private Const READY_STABLE_MS As Long = 500
Private Const READY_POLL_MS As Long = 200

Private Type TThis
    wv As stdWebView
    site As String
    siteBase As String
    readyTimeoutMs As Long
    isSessionReady As Boolean
End Type
Private This As TThis

'Create a lazy SharePoint authenticator.
'No WebView is shown or initialized until protEnsureAuthenticated/stdICallable_Run is called.
'@param SiteUrl - The SharePoint site URL to authenticate to.
'@param ReadyTimeoutMs - The maximum time to wait for the SharePoint site to be ready.
'@returns - A stdSharepointAuthenticator instance.
Public Function Create( _
    ByVal SiteUrl As String, _
    Optional ByVal ReadyTimeoutMs As Long = DEFAULT_READY_TIMEOUT_MS) As stdSharepointAuthenticator

    Dim instance As stdSharepointAuthenticator
    Set instance = New stdSharepointAuthenticator
    Call instance.protInit(SiteUrl, ReadyTimeoutMs)
    Set Create = instance
End Function

'Initialize lightweight config only (lazy startup).
'@protected
'@param SiteUrl - The SharePoint site URL to authenticate to.
'@param ReadyTimeoutMs - The maximum time to wait for the SharePoint site to be ready.
Public Sub protInit( _
    ByVal SiteUrl As String, _
    Optional ByVal ReadyTimeoutMs As Long = DEFAULT_READY_TIMEOUT_MS)

    If LenB(Trim$(SiteUrl)) = 0 Then
        Err.Raise 5, "stdSharepointAuthenticator::protInit", "SiteUrl cannot be blank"
    End If
    This.site = Trim$(SiteUrl)
    This.siteBase = InferSiteBase(This.site)
    This.readyTimeoutMs = IIf(ReadyTimeoutMs > 0, ReadyTimeoutMs, DEFAULT_READY_TIMEOUT_MS)
    This.isSessionReady = False
End Sub

'Force authentication upfront (optional eager path).
'@protected
Public Sub protEnsureAuthenticated()
    If This.wv Is Nothing Then Set This.wv = stdWebView.CreateFromUserform(Me)
    Call Me.Show(False)
    Call This.wv.Navigate(This.site)
    If Not This.wv.WaitForDocumentReady(This.siteBase, This.ReadyTimeoutMs, READY_STABLE_MS, READY_POLL_MS) Then
        Me.Hide
        Err.Raise 5, "stdSharepointAuthenticator::protEnsureAuthenticated", "Timed out waiting for SharePoint authentication"
    End If
    This.isSessionReady = True
    Me.Hide
End Sub

'Release hosted WebView resources.
'@returns - None.
Public Sub Quit()
    On Error Resume Next
    If Not This.wv Is Nothing Then This.wv.Quit
    Set This.wv = Nothing
    This.isSessionReady = False
End Sub

'Release hosted WebView resources on form terminate.
Private Sub UserForm_Terminate()
    On Error Resume Next
    Call Quit
End Sub

'Build a cookie header for a given target URL.
'@param targetUrl - The URL to build a cookie header for.
'@returns - A string containing the cookie header.
Private Function BuildCookieHeader(ByVal targetUrl As String) As String
    Dim pairs As Variant
    Dim i As Long
    Dim part As String
    Dim headerText As String

    pairs = This.wv.CookiesForRequest(targetUrl)
    If Not IsArray(pairs) Then Exit Function
    On Error GoTo noCookies
    For i = LBound(pairs) To UBound(pairs) Step 2
        If i + 1 > UBound(pairs) Then Exit For
        part = CStr(pairs(i)) & "=" & CStr(pairs(i + 1))
        If LenB(headerText) = 0 Then
            headerText = part
        Else
            headerText = headerText & "; " & part
        End If
    Next i
noCookies:
    BuildCookieHeader = headerText
End Function

'Resolve the target URL for a given request URL.
'@param requestUrl - The URL to resolve the target URL for.
'@returns - The target URL.
Private Function ResolveTargetUrl(ByVal requestUrl As String) As String
    requestUrl = Trim$(requestUrl)
    If LenB(requestUrl) > 0 Then
        ResolveTargetUrl = requestUrl
    Else
        ResolveTargetUrl = This.site
    End If
End Function

'Run the authenticator.
'@param params - The parameters to pass to the authenticator.
'@returns - The result of the authenticator.
Private Function stdICallable_Run(ParamArray params() As Variant) As Variant
    stdICallable_Run = stdICallable_RunEx(params)
End Function

'Run the authenticator with explicit parameters.
'@param params - The parameters to pass to the authenticator.
'@returns - The result of the authenticator.
Private Function stdICallable_RunEx(ByVal params As Variant) As Variant
    Dim pHTTP As Object
    Dim requestUrl As String
    Dim targetUrl As String
    Dim cookieHeader As String

    If Not IsArray(params) Then
        Err.Raise 5, "stdSharepointAuthenticator::stdICallable_RunEx", "Expected parameter array"
    End If
    If UBound(params) < 2 Then
        Err.Raise 5, "stdSharepointAuthenticator::stdICallable_RunEx", "Expected (pHTTP, RequestMethod, sURL, ...)"
    End If

    Set pHTTP = params(0)
    requestUrl = CStr(params(2))
    targetUrl = ResolveTargetUrl(requestUrl)
    If LenB(targetUrl) = 0 Then
        Err.Raise 5, "stdSharepointAuthenticator::stdICallable_RunEx", "No request URL available for authentication"
    End If

    If Not This.isSessionReady Then
        Call protEnsureAuthenticated()
    End If

    cookieHeader = BuildCookieHeader(targetUrl)
    If LenB(cookieHeader) = 0 Then
        Err.Raise 5, "stdSharepointAuthenticator::stdICallable_RunEx", "No cookies available for URL: " & targetUrl
    End If

    Call pHTTP.SetRequestHeader("Cookie", cookieHeader)
End Function

'Bind the authenticator to a set of parameters.
'@param params - The parameters to bind to the authenticator.
'@returns - The bound authenticator.
Private Function stdICallable_Bind(ParamArray params() As Variant) As stdICallable
    Err.Raise 5, "stdSharepointAuthenticator::stdICallable_Bind", "Bind is not implemented"
End Function

'Send a message to the authenticator.
'@param sMessage - The message to send to the authenticator.
'@param success - Whether the message was successfully sent.
'@param params - The parameters to send to the authenticator.
'@returns - The result of the message.
Private Function stdICallable_SendMessage(ByVal sMessage As String, ByRef success As Boolean, ByVal params As Variant) As Variant
    Select Case LCase$(Trim$(sMessage))
        Case "classname"
            success = True
            stdICallable_SendMessage = "stdSharepointAuthenticator"
        Case "obj"
            success = True
            Set stdICallable_SendMessage = Me
        Case Else
            success = False
    End Select
End Function

'Infer the site base from a given raw URL.
'@param rawUrl - The raw URL to infer the site base from.
'@returns - The site base.
Private Function InferSiteBase(ByVal rawUrl As String) As String
    Dim s As String
    Dim protocolEnd As Long
    Dim pathStart As Long
    Dim hostPart As String
    Dim pathPart As String
    Dim firstSeg As String
    Dim secondSlash As Long

    s = Trim$(rawUrl)
    If LenB(s) = 0 Then Exit Function

    protocolEnd = InStr(1, s, "://", vbTextCompare)
    If protocolEnd <= 0 Then
        InferSiteBase = s
        Exit Function
    End If

    pathStart = InStr(protocolEnd + 3, s, "/")
    If pathStart <= 0 Then
        InferSiteBase = s
        Exit Function
    End If

    hostPart = Left$(s, pathStart - 1)
    pathPart = Mid$(s, pathStart)
    If Left$(LCase$(pathPart), 7) = "/sites/" Or Left$(LCase$(pathPart), 7) = "/teams/" Then
        secondSlash = InStr(8, pathPart, "/")
        If secondSlash > 0 Then
            firstSeg = Left$(pathPart, secondSlash - 1)
        Else
            firstSeg = pathPart
        End If
        InferSiteBase = hostPart & firstSeg
    Else
        InferSiteBase = hostPart
    End If
End Function
