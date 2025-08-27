VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ClipboardInspector 
   Caption         =   "Clipboard Inspector"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11400
   OleObjectBlob   =   "ClipboardInspector.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ClipboardInspector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TThis
  ClipboardID As Long
  Form As stdUIElement
  Formats As stdUIElement
  Frame As stdUIElement
  Anchors(1 To 3) As TElementAnchor
End Type
Private This As TThis
Private WithEvents timer As stdTimer
Attribute timer.VB_VarHelpID = -1





Private Sub CommandButton1_Click()
  Call timer_Tick
End Sub


Private Sub lvClipboardFormats_ItemClick(ByVal item As MSComctlLib.ListItem)
  Call frPreview.Controls.Clear
  
  Dim Element As stdUIElement
  Select Case CLng(item.SubItems(1))
    Case CF_BITMAP, CF_ENHMETAFILE
      Set Element = stdUIElement.CreateFromType(frPreview.Controls, uiImage)
      Dim pic As MSForms.Image: Set pic = Element.UIObject
      Set pic.Picture = stdClipboard.Picture
      With Element
        .Width = frPreview.Width
        .Height = frPreview.Height
        With .UIObject
          .PictureSizeMode = MSForms.fmPictureSizeModeZoom
        End With
      End With
      
    Case CF_TEXT, CF_UNICODETEXT
      Set Element = stdUIElement.CreateFromType(frPreview.Controls, uiTextBox, value:=stdClipboard.text)
      With Element
        .Width = frPreview.Width
        .Height = frPreview.Height
        With .UIObject
          .FontName = "Consolas"
          .WordWrap = True
          .MultiLine = True
        End With
      End With
      
    Case Else
      Dim bytes: bytes = stdClipboard.value(CLng(item.SubItems(1)))
      
      'hexdump bytes
      Dim hexDump As String: hexDump = String(3 * LenB(bytes) - 1, " ")
      For i = 0 To LenB(bytes) - 1
        Mid(hexDump, 3 * i + 1, 2) = IIf(bytes(i + 1) < 16, "0", "") & Hex(bytes(i + 1))
      Next
      
      'Serialize with Unicode/Ascii prediction
      Dim stringified As String
      If bytes(2) = 0 Then
        stringified = bytes
      Else
        stringified = StrConv(bytes, vbUnicode)
      End If
      stringified = Replace(stringified, vbNullChar, "")
      
      'Create element and set settings
      Dim finalText As String
      finalText = stringified & vbCrLf & vbCrLf & "-----------------------------------" & vbCrLf & vbCrLf & hexDump
      Set Element = stdUIElement.CreateFromType(frPreview.Controls, uiTextBox, value:=finalText)
      Dim tb As MSForms.TextBox
      
      With Element
        .Width = frPreview.Width
        .Height = frPreview.Height
        With .UIObject
          .FontName = "Consolas"
          .WordWrap = True
          .MultiLine = True
          .ScrollBars = fmScrollBarsVertical
        End With
        .value = finalText
      End With
  End Select
  
  This.Anchors(3) = TElementAnchor_Create(frPreview, Element, AnchorTypeFixed, AnchorTypeFixed, AnchorTypePercentile, AnchorTypePercentile)
End Sub

Private Sub timer_Tick()
  Dim latestClipboardID As Long: latestClipboardID = stdClipboard.ClipboardID
  If latestClipboardID <> This.ClipboardID Then
    This.ClipboardID = latestClipboardID
    
    lvClipboardFormats.ListItems.Clear
    Dim ids As Collection: Set ids = stdClipboard.formatIDs
    Dim Formats: Set Formats = stdClipboard.Formats
    Dim i As Long
    For i = 1 To ids.Count
      If stdClipboard.IsFormatAvailable(ids(i)) Then
        Dim li As ListItem
        Set li = lvClipboardFormats.ListItems.Add(text:=format(ids(i), "00000#"))
        Call li.ListSubItems.Add(text:=ids(i))
        Call li.ListSubItems.Add(text:=Formats(i))
        Call li.ListSubItems.Add(text:=stdClipboard.FormatSize(ids(i)))
      Else
        Call GlobalLog("ClipboardInspector", "  Clipboard format not available: " & ids(i) & ", " & Formats(i))
      End If
    Next
  End If
End Sub



Private Sub UserForm_Initialize()
  Set timer = stdTimer.Create(300)
  With lvClipboardFormats.ColumnHeaders.Add(text:="SortKey")
    .Width = 0
  End With
  lvClipboardFormats.ColumnHeaders.Add text:="ID"
  lvClipboardFormats.ColumnHeaders.Add text:="Name"
  lvClipboardFormats.ColumnHeaders.Add text:="Size"
  With stdWindow.CreateFromIUnknown(Me)
    .isResizable = True
    .isAlwaysOnTop = True
    Call .setOwnerHandle(0)
  End With
  
  This.Anchors(1) = TElementAnchor_Create(Me, lvClipboardFormats, AnchorTypeFixed, AnchorTypeFixed, AnchorTypePercentile, AnchorTypePercentile)
  This.Anchors(2) = TElementAnchor_Create(Me, frPreview, AnchorTypeFixed, AnchorTypePercentile, AnchorTypePercentile, AnchorTypePercentile)
End Sub

Private Sub UserForm_Resize()
  Dim i As Long
  For i = 1 To 3
    If Not This.Anchors(i).Element Is Nothing Then
      Call TElementAnchor_Resize(This.Anchors(i))
    End If
  Next
End Sub


