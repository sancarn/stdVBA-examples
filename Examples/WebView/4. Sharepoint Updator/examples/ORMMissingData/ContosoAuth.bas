Attribute VB_Name = "ContosoAuth"

Private auth As SharepointAuthenticator
Public Property Get Authenticator() As SharepointAuthenticator
  If auth Is Nothing Then Set auth = SharepointAuthenticator.Create("https://contoso.sharepoint.com")
  Set Authenticator = auth
End Property
