VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sharepoint 
   Caption         =   "Sharepoint"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   OleObjectBlob   =   "Sharepoint.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SharepointList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type TThis
   wv as stdWebView
   site as string
End Type
Private This As TThis

'Creates a userform for updating a sharepoint list.
'@constructor
'@param site - The sharepoint site url.
'@returns - Sharepoint list userform.
'@example `set sharepoint = Sharepoint.Create("https://contoso.sharepoint.com/sites/MySite/Lists/MyList/AllItems.aspx")`
Public Function Create(ByVal site As String) As SharepointList
  Set Create = New SharepointList
  Call Create.protInit(site)
  Unload Me
End Function

'Initialise the userform
'@protected
'@param site - The sharepoint site url.
Public Sub protInit(ByVal site As String)
  Set This.wv = stdWebView.CreateFromUserform(Me)
  This.site = site
  Call This.wv.Navigate(This.site)
  Call Me.Show(False)
End Sub