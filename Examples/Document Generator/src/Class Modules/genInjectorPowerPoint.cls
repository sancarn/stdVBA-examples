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

#If VBA7 Then
  Private Declare PtrSafe Function PostMessageA Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Any, ByVal lParam As Any) As Long
  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
  Private Declare Function PostMessageA Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Any, ByVal lParam As Any) As Long
  Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
#End If

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
  Dim bindings As Collection: Set bindings = FindAll(This.templatePresentation, "*$*")
  Dim tmp As Collection: Set tmp = New Collection
  For Each shape In bindings
    Dim address As String: address = shape.Parent.name & "::" & shape.name
    Call tmp.add(CreateDictionary( _
      "lambda", shape.textFrame.TextRange.Text, _
      "setter", stdCallback.CreateFromObjectMethod(Me, "setValue").Bind(address), _
      "getSetterTarget", stdCallback.CreateFromObjectMethod(Me, "getShape").Bind(address), _
      "meta", CreateDictionary() _
    ))
  Next

  Set getFormulaBindings = stdEnumerator.CreateFromIEnumVariant(tmp)
End Function


'Initialise the target presentation
Public Function InitialiseTarget() As Object
  Set This.targetPresentation = This.appPowerPoint.Presentations.add()
  'this.targetPresentation.ApplyTemplate(This.templatePath) '<== Not sure if this works for non .potx files
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
  Dim shape As Object: Set shape = getTargetBinding(address)
  Dim slide As Object: Set slide = shape.Parent
  If IsObject(value) Then
    Select Case TypeName(value)
      Case "Dictionary"
        If value.Exists("Type") Then
          Select Case value("Type")
            Case "Picture"
              Dim p As Object: Set p = slide.shapes.AddPicture(value("Path"), msoTrue, msoTrue, 0, 0)
              Call shape.Fill.UserPicture(value("Path"))
              Dim factor As Double: factor = min(shape.height / p.height, shape.width / p.width)
              shape.PictureFormat.Crop.PictureWidth = p.width * factor
              shape.PictureFormat.Crop.PictureHeight = p.height * factor
              p.Delete
              shape.textFrame.TextRange.Text = ""
            Case Else
              shape.textFrame.TextRange.Text = "ERROR: Unsupported dictionary type '" & value("Type") & "'"
          End Select
        Else
          shape.textFrame.TextRange.Text = "ERROR: Unsupported dictionary"
        End If
      Case Else
          shape.textFrame.TextRange.Text = "ERROR: Unsupported object type '" & TypeName(value) & "'"
    End Select
  Else
      shape.textFrame.TextRange.Text = value
  End If
End Sub

Public Function getShape(ByVal address As String) As Object
  Set getShape = getTargetBinding(address)
End Function

'Obtain a binding from the target
'@param address - The address of a shape to get
'@returns Object<PowerPoint.Shape> - Powerpoint shape object for the current target
Private Function getTargetBinding(ByVal address As String) As Object
  Dim slideName As String: slideName = Split(address, "::")(0)
  Dim shapeName As String: shapeName = Split(address, "::")(1)
  Dim slide As Object: Set slide = This.targetPresentation.Slides(slideName)
  Set getTargetBinding = slide.shapes(shapeName)
End Function

'Find all the shapes in the presentation that contain text
'@param presentation - The presentation to search
'@param findWhat - The text to search for
'@returns Collection<Object<Powerpoint.Shape>> - A collection of shapes that contain the text
Private Function FindAll(ByVal presentation As Object, ByVal findWhat As String) As Collection
  Dim c As Collection: Set c = New Collection

  Dim slide As Object
  For Each slide In presentation.Slides
    Dim shape As Object
    For Each shape In slide.shapes
      If shape.HasTextFrame Then
        Dim textFrame As Object: Set textFrame = shape.textFrame
        If textFrame.HasText Then
          Dim Text As String: Text = textFrame.TextRange.Text
          If Text Like "*$*" Then
            Call c.add(shape)
          End If
        End If
      End If
    Next
  Next
  Set FindAll = c
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


