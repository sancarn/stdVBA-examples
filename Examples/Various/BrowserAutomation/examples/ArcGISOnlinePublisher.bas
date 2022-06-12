Attribute VB_Name = "ArcGISOnlinePublisher"

'Edit accordingly:
Const AGISOSERVER  = "SERVER.maps.arcgis.com"
Const TestLayerCSV = "MY_CSV_LOCATION"
Const TestLayerURL = "https://$SERVER/home/item.html?id=MY_LAYER_ID"
Const SPOCHECKER   = "*SERVER.sharepoint.com*"


Public Sub test()
    Call PublishCSV(SI(TestLayerURL), TestLayerCSV)
End Sub

Public Sub PublishCSV(ByVal sLayerURL As String, ByVal sCSVPath As String)
    'Force CSV to be local ready for uploading to GISSTOnline
    Dim sCSV As String: sCSV = SPOToCSV(sCSVPath)
    
    'Input username and password for ArcGIS Online
    Dim C_USER as string: C_USER = InputBox("Enter your username")
    Dim C_PASS as string: C_PASS = InputBox("Enter your password")

    'Launch chrome
    Dim chrome As stdChrome: Set chrome = stdChrome.Create()
    
    'Navigate to ArcGIS login page
    Call chrome.Navigate(SI("https://$SERVER/home/signin.html?useLandingPage=true"))
    Call chrome.AwaitForCondition(stdLambda.Create("$2.address like $1").Bind(SI("$SERVER/sharing/*oauth2/authorize?client_id=arcgisonline*")))
    
    'Login and wait till login authorised
    Dim accLogin As stdAcc: Set accLogin = chrome.accMain.AwaitForElement(stdLambda.Create("$1.name = ""ArcGIS login"" and $1.role = ""ROLE_PANE"""))
    Dim accUser As stdAcc: Set accUser = accLogin.FindFirst(stdLambda.Create("$1.Name = ""Username"" and $1.Role = ""ROLE_TEXT"""))
    Dim accPass As stdAcc: Set accPass = accLogin.FindFirst(stdLambda.Create("$1.Name = ""Password"" and $1.Role = ""ROLE_TEXT"""))
    Dim accSignIn As stdAcc: Set accSignIn = accLogin.FindFirst(stdLambda.Create("$1.Name = ""Sign In"" and $1.Role = ""ROLE_PUSHBUTTON"""))
    accUser.value = C_USER
    accPass.value = C_PASS
    Call accSignIn.DoDefaultAction
    Call chrome.AwaitForCondition(stdLambda.Create("$2.Address like $1").Bind(SI("$SERVER/home/index.html")))
    
    'Navigate to layer url
    Call chrome.Navigate(sLayerURL)
    'HACK: This is quite slow, consider attempting to speed this up with more tree refinement
    
    'Await path "4.1.1.2.2.2.4.2", we do this somewhat dynamically though
    Call chrome.AwaitForCondition(stdLambda.Create("$1.winMain.Caption like ""* - Overview - *"""))
    Dim sDocCaption As String: sDocCaption = left(chrome.winMain.Caption, Len(chrome.winMain.Caption) - 16)
    Dim accDoc As stdAcc: Set accDoc = chrome.AwaitForAccElement(stdLambda.Create("$2.Name = $1 and $2.Role = ""ROLE_DOCUMENT""").Bind(sDocCaption))
    
    'Wait til no longer loading and property page is visible
    Dim accMenu As stdAcc: Set accMenu = accDoc.CreateFromPath("4.2")
    Do
        Set accMenu = accDoc.CreateFromPath("4.2")
        DoEvents
    Loop Until accMenu.Role = "ROLE_PROPERTYPAGE"
    
    'Click update data and overwrite layer
    Call accMenu.AwaitForElement(stdLambda.Create("$1.Name = ""Update Data"" and $1.Role = ""ROLE_PUSHBUTTON""")).DoDefaultAction               '$1.Role = ""
    Call accMenu.AwaitForElement(stdLambda.Create("$1.Name = ""Overwrite Entire Layer"" and $1.Role = ""ROLE_PUSHBUTTON""")).DoDefaultAction
    
    'Click Choose file button
    Call accDoc.AwaitForElement(stdLambda.Create("$1.Name = ""Choose file"" and $1.Role = ""ROLE_PUSHBUTTON""")).DoDefaultAction
    
    'Await the file uploader and target the CSV for uploading 'FIX: Check windw exists to better handle errors
    Set fileUploader = stdWindow.CreateFromDesktop().AwaitForWindow(stdLambda.Create("if $2 <= 1 and $1.exists then $1.Class = ""#32770"" and $1.Caption = ""Open"" else EWndFindResult.NoMatchSkipDescendents"))
    Dim accFileUploader As stdAcc: Set accFileUploader = stdAcc.CreateFromHwnd(fileUploader.Handle)
    accFileUploader.FindFirst(stdLambda.Create("$1.Role = ""ROLE_TEXT"" and $1.Name = ""File name:""")).value = sCSV
    Call accFileUploader.FindFirst(stdLambda.Create("$1.DefaultAction = ""Press"" and $1.Name = ""Open""")).DoDefaultAction
    
    'Click overwrite button
    Call accDoc.AwaitForElement(stdLambda.Create("$1.Name = ""Overwrite"" and $1.Role = ""ROLE_PUSHBUTTON""")).DoDefaultAction
    Call chrome.AwaitForCondition(stdLambda.Create("$1.Address like ""*&jobid=*"""))
    
    'Quit chrome
    Call chrome.Quit
End Sub

'
Private Function SI(ByVal s as string) as string
    SI = replace(s, "$SERVER", AGISOSERVER)
End Function

'
Private Function SPOToCSV(ByVal sSharepointOnlineURL) As String
    'If CSVPath on sharepoint then we need to make a local copy, because windows explorer isn't able to use sharepoint directly
    If LCase(sSharepointOnlineURL) Like SPOCHECKER Then
        With Workbooks.Open(sSharepointOnlineURL)
            Application.DisplayAlerts = False
                Dim sFileName As String: sFileName = FileNameFromURL(sSharepointOnlineURL)
                .SaveAs "C:\Temp\" & sFileName
                SPOToCSV = "C:\Temp\" & sFileName
                .Close False
            Application.DisplayAlerts = True
        End With
    Else
        SPOToCSV = sSharepointOnlineURL
    End If
End Function

'Get file name from url
'@param {String} URL to get file name from
'@returns {String} File name from url
Private Function FileNameFromURL(ByVal sURL As String) As String
    Dim sFileName As String
    sFileName = Mid(sURL, InStrRev(sURL, "/") + 1)
    sFileName = Replace(sFileName, "%20", " ")
    FileNameFromURL = sFileName
End Function
