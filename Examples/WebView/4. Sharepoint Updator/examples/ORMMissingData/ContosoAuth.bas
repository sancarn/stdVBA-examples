Attribute VB_Name = "ContosoAuth"

Private auth As stdSharepointAuthenticator
Public Property Get Authenticator() As stdSharepointAuthenticator
  If auth Is Nothing Then Set auth = stdSharepointAuthenticator.Create("https://contoso.sharepoint.com")
  Set Authenticator = auth
End Property
