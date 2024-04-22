VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "genInjectorPowerPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements genIInjector

Private Type TThis
  templatePath As String
  templatePresentation As Object 'PowerPoint.Presentation
  appPowerPoint As Object      'PowerPoint.Application
  targetPresentation As Object 'PowerPoint.Presentation
End Type
Private This As TThis

'Create a new instance of the injector
'@param templatePath - The path to the template file
'@returns - An instance of a genIInjector object - within here include all initialisation stuff
Public Function Create(ByVal templatePath As String) As genIInjector
  Dim gen As New genInjectorPowerPoint
  Call gen.protInit(templatePath)
  Set Create = gen
End Function

'Initialise the injector
'@param templatePath - The path to the template file
'@param templateSheet - The name of the sheet in the template file
Public Sub protInit(ByVal templatePath As String)
  This.templatePath = templatePath
  Set This.appPowerPoint = CreateObject("PowerPoint.Application")
  This.appPowerPoint.Visible = True
  Set This.templatePresentation = This.appPowerPoint.Presentations.open(This.templatePath) '(For use in getFormulaBindings)
End Sub

'Obtain the bindings for the injector
'@returns stdEnumerator<Object<Dictionary<"lambda", string, "setter", stdICallable>>> - a list of bindings for the injector
'@example ```
'  For each binding in injector.getFormulaBindings()
'    binding("setter").call(stdLambda.create(binding("lambda")).call(row))
'  Next
'```
Public Function getFormulaBindings() As stdEnumerator
  Dim tmp As Collection: Set tmp = New Collection

  Dim slide As Object
  For Each slide In This.templatePresentation.Slides
    'If notes present then add to bindings
    If slide.HasNotesPage Then
      Dim notes As String: notes = slide.NotesPage.Shapes(2).textFrame.TextRange.Text
      If notes Like "*$*" Then
        Dim address As String: address = "slide::" & slide.name
        Call tmp.add(CreateDictionary( _
          "lambda", notes, _
          "setter", stdCallback.CreateFromObjectMethod(Me, "setValue").Bind(address), _
          "getSetterTarget", stdCallback.CreateFromObjectMethod(Me, "getTargetBinding").Bind(address), _
          "meta", CreateDictionary() _
        ))
      End If
    End If

    Dim shape As Object
    For Each shape In slide.Shapes
      Call handleBindingsForShape(slide, shape, tmp)
    Next
  Next

  Set getFormulaBindings = stdEnumerator.CreateFromIEnumVariant(tmp)
End Function

'Append all bindings for a given shape to the binding collection
'@param slide - The slide the shape is on
'@param shape - The shape to handle the bindings for
'@param bindings - The collection to add the bindings to
Private Sub handleBindingsForShape(ByVal slide as object, ByVal shape as object, ByVal bindings as Collection)
  const msoGroup = 6
  const msoTable = 19
  Dim Text As String, address As String
  select case shape.type
    case msoGroup
      Dim subShape as Object
      For Each subShape In shape.GroupItems
        Call handleBindingsForShape(slide, subShape, bindings)
      Next
    case msoTable
      Dim row as Object, iRow as long: iRow = 0
      For Each row In shape.table.Rows
        iRow = iRow + 1
        Dim cell as Object, iCell as long: iCell = 0
        For Each cell In row.Cells
          iCell = iCell + 1
          Text = cell.Shape.textFrame.TextRange.Text
          If Text Like "*$*" Then
            address = "table::" & slide.name & "::" & shape.name & "::" & iRow & "::" & iCell
            Call bindings.add(CreateDictionary( _
              "lambda", Text, _
              "setter", stdCallback.CreateFromObjectMethod(Me, "setValue").Bind(address), _
              "getSetterTarget", stdCallback.CreateFromObjectMethod(Me, "getTargetBinding").Bind(address), _
              "meta", CreateDictionary() _
            ))
          End If
        Next
      Next
    case else
      If shape.HasTextFrame Then
        If shape.textFrame.HasText Then
          Text = shape.textFrame.TextRange.Text
          If Text Like "*$*" Then
            address = "shape::" & slide.name & "::" & shape.name
            Call bindings.add(CreateDictionary( _
              "lambda", Text, _
              "setter", stdCallback.CreateFromObjectMethod(Me, "setValue").Bind(address), _
              "getSetterTarget", stdCallback.CreateFromObjectMethod(Me, "getTargetBinding").Bind(address), _
              "meta", CreateDictionary() _
            ))
          End If
        End If
      End If
  end select
End Sub


'Initialise the target presentation
Public Function InitialiseTarget() As Object
  'Create presentation
  Set This.targetPresentation = This.appPowerPoint.Presentations.add()
  
  'Ensure slide layout is the same as the template
  This.targetPresentation.PageSetup.SlideOrientation = This.templatePresentation.PageSetup.SlideOrientation
  This.targetPresentation.PageSetup.SlideWidth = This.templatePresentation.PageSetup.SlideWidth
  This.targetPresentation.PageSetup.SlideHeight = This.templatePresentation.PageSetup.SlideHeight
  
  'Insert slides from template
  Call This.targetPresentation.Slides.insertFromFile(This.templatePath, 0)
  
  'Return presentation object
  Set InitialiseTarget = This.targetPresentation
End Function

'Cleanup the target workbook
Public Sub CleanupTarget()
  Call ClosePresentation(This.targetPresentation)
End Sub

'Delegate for "setter" property
'@param address - The address of the shape to set the value of (in the form "SlideName::ShapeName")
'@param value - The value to set the shape content to
Public Sub setValue(ByVal address As String, ByVal value As Variant)
  Const ppAlertsAll As Long = 2
  Const ppAlertsNone As Long = 1
  Dim targetType As String: targetType = Split(address, "::")(0)
  Dim target As Object: Set target = getTargetBinding(address)
  Dim slide As Object: Set slide = getTargetBindingSlide(address)
  If IsObject(value) Then
    Select Case TypeName(value)
      Case "Dictionary"
        If value.Exists("Type") Then
          Select Case value("Type")
            Case "Picture"
              select case targetType 
                case "shape", "table"
                  On Error GoTo ErrorImportingImage
                  This.appPowerPoint.DisplayAlerts = ppAlertsNone
                  Dim p As Object: Set p = slide.Shapes.AddPicture(value("Path"), msoTrue, msoTrue, 0, 0)
                  This.appPowerPoint.DisplayAlerts = ppAlertsAll
                  On Error GoTo 0
                  Call target.Fill.UserPicture(value("Path"))
                  Dim factor As Double: factor = min(target.height / p.height, target.width / p.width)
                  target.PictureFormat.Crop.PictureWidth = p.width * factor
                  target.PictureFormat.Crop.PictureHeight = p.height * factor
                  p.Delete
                  target.textFrame.TextRange.Text = ""
                  Exit Sub
ErrorImportingImage:
                  target.textFrame.TextRange.Text = "ERROR: Picture cannot be imported."
                  Exit Sub
                case else
                  Err.Raise 1, "setValue", "ERROR: Pictures can only be added to shapes, not " & targetType & "s"
              End Select
            Case Else
              Select Case targetType
                Case "shape", "table"
                  target.textFrame.TextRange.Text = "ERROR: Unsupported dictionary type '" & value("Type") & "'"
                Case Else
                  Err.Raise 1, "setValue", "ERROR: Unsupported dictionary type '" & value("Type") & "'"
              End Select
          End Select
        Else
          Select Case targetType
            Case "shape", "table"
              target.textFrame.TextRange.Text = "ERROR: Unsupported dictionary"
            Case Else
              Err.Raise 1, "setValue", "ERROR: Unsupported dictionary"
          End Select
        End If
      Case Else
        Select Case targetType
          Case "shape", "table"
            target.textFrame.TextRange.Text = "ERROR: Unsupported object type '" & TypeName(value) & "'"
          Case Else
            Err.Raise 1, "setValue", "ERROR: Unsupported object type '" & TypeName(value) & "'"
        End Select
    End Select
  Else
    Select Case targetType
      Case "shape", "table"
        target.textFrame.TextRange.Text = value
      Case Else
        'Do nothing
    End Select
  End If
End Sub

'Obtain a binding from the target
'@param address - The address of a shape to get
'@returns Object<PowerPoint.Shape> - Powerpoint shape object for the current target
Public Function getTargetBinding(ByVal address As String) As Object
  Dim addressParts As Variant: addressParts = Split(address, "::")
  Select Case addressParts(0)
    Case "slide"
      Set getTargetBinding = This.targetPresentation.Slides(addressParts(1))
    Case "shape"
      Set getTargetBinding = This.targetPresentation.Slides(addressParts(1)).Shapes(addressParts(2))
    case "table"
      Set getTargetBinding = This.targetPresentation.Slides(addressParts(1)).Shapes(addressParts(2)).table.cell(addressParts(3), addressParts(4)).Shape
  End Select
End Function

'Obtain a binding slide from the target
'@param address - The address of a binding to get the slide for
'@returns Object<PowerPoint.Slide> - Powerpoint slide object for the current target
Public Function getTargetBindingSlide(ByVal address as string) as Object
  Dim addressParts As Variant: addressParts = Split(address, "::")
  set getTargetBindingSlide = This.targetPresentation.Slides(addressParts(1))
End Function

'Find the minimum of two numbers
'@param a - The first number
'@param b - The second number
'@returns - The minimum of the two numbers
Private Function min(ByVal a As Double, ByVal b As Double) As Double
  If a < b Then
    min = a
  Else
    min = b
  End If
End Function

'Create a dictionary from a list of key-value pairs
'@param v - A list of key-value pairs
'@returns Object<Dictionary> - A dictionary object
Private Function CreateDictionary(ParamArray v() As Variant) As Object
  Set CreateDictionary = CreateObject("Scripting.Dictionary")
  Dim i As Long
  For i = 0 To UBound(v) Step 2
    CreateDictionary.add v(i), v(i + 1)
  Next
End Function

'This function is implemented as a fix for an issue where VBA reports "Presentation.Close: Failed"
'@param pres as Object<PowerPoint.Presentation> - The presentation to close.
'@docs https://stackoverflow.com/questions/78156015/vba-powerpoint-run-time-error-2147467259-80004005-presentation-close-fail/78156016#78156016
Public Sub ClosePresentation(ByVal pres As Object)
  pres.windows(1).Activate
  Dim i As Long
  While Not pres.Saved
    i = i + 1
    pres.Saved = True
    Debug.Print i
  Wend
  stdWindow.CreateFromApplication(pres.Application).Quit
End Sub

'**********
'* Events *
'**********

Private Sub Class_Terminate()
  Call ClosePresentation(This.templatePresentation)
  This.appPowerPoint.Quit
End Sub

Private Function genIInjector_Create(ByVal templatePath As String, ParamArray v() As Variant) As genIInjector
  Set genIInjector_Create = Create(templatePath)
End Function
Private Function genIInjector_getFormulaBindings() As stdEnumerator
  Set genIInjector_getFormulaBindings = getFormulaBindings()
End Function
Private Function genIInjector_InitialiseTarget() As Object
  Set genIInjector_InitialiseTarget = InitialiseTarget()
End Function
Private Sub genIInjector_CleanupTarget()
  Call CleanupTarget
End Sub




