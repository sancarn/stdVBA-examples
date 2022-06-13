Public Enum BrowsersEnum
    Brave
    Chrome
    Edge
    FireFox
    InternetExplorer
    Opera
    Vivaldi
End Enum

Public Sub SideBySide(Optional SelectedBrowser As BrowsersEnum = Chrome)
  'Identify different browsers by Window Class and Caption pattern matcher
  Dim sBrowserCaptionMatcher as string, sBrowserClass as string
  Select case SelectedBrowser
    case Brave
      sBrowserClass = "Chrome_WidgetWin_1"
      sBrowserCaptionMatcher = "*- Brave"
    case Chrome
      sBrowserClass = "Chrome_WidgetWin_1"
      sBrowserCaptionMatcher = "*- Google Chrome"
    case Edge
      sBrowserClass = "Chrome_WidgetWin_1"
      sBrowserCaptionMatcher = "*- Microsoft Edge"
    case FireFox
      sBrowserClass = "MozillaWindowClass"
      sBrowserCaptionMatcher = "*- Mozilla FireFox"
    Case InternetExplorer
      sBrowserClass = "IEFrame"
      sBrowserCaptionMatcher = "*- Internet Explorer"
    Case Opera
      sBrowserClass = "Chrome_WidgetWin_1"
      sBrowserCaptionMatcher = "*- Opera"
    Case Vivaldi
      sBrowserClass = "Chrome_WidgetWin_1"
      sBrowserCaptionMatcher = "*- Vivaldi"
  end select

  'Create a matcher for FindFirst
  Dim browserFinder as stdLambda
  set browserFinder = stdLambda.Create("if $4 > 1 then EWndFindResult.NoMatchSkipDescendents else $3.Class = $2 And $3.Caption like $1") _ 
    .bind(sBrowserCaptionMatcher, sBrowserClass)
  
  'Get the desktop. Set  the application to the left half of the screen, set the web browser to the right half of the screen.
  Dim desktop As stdWindow: Set desktop = stdWindow.CreateFromDesktop()
  With stdWindow.CreateFromApplication()
    .State = Normal
    .x = desktop.x
    .y = desktop.y
    .width = desktop.width / 2
    .height = desktop.height - 40
  End With
  
  Dim win as stdWindow
  For each win in stdWindow.CreateFromDesktop.FindAll(browserFinder)
    .State = Normal
    .x = desktop.x + desktop.width / 2
    .y = desktop.y
    .width = desktop.width / 2
    .height = desktop.height - 40
    .Activate
  next
End Sub
