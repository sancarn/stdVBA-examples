# wuiUI

`wuiUI` is a VBA-first UI framework concept that combines:

- `stdHTML` for declarative UI tree construction and traversal.
- `stdWebView` for rendering, hosting, and browser interop.
- `stdICallable` for callback/event plumbing.

The goal is a React-like developer experience (declarative components, predictable updates), while keeping UI state and update logic on the VBA side.

## Design Intent

- VBA is the source of truth for UI state.
- Rendering can happen in WebView, but state transitions are driven by VBA.
- UI updates are deterministic and explicit (not hidden in large JS frameworks).
- Callbacks use `stdICallable`-compatible objects for composable event handling.
- Every element gets an internal stable ID so callbacks can safely target nodes/state.

## Current Building Blocks

### `stdHTML` (available)

- Fluent API for building element trees (`CreateChild`, `CreateLiteral`, `Raw`).
- Query/update helpers (`QuerySelector`, `QuerySelectorAll`, `ReplaceWith`).
- Suitable as the core virtual/UI tree model managed by VBA.

### `stdWebView` (available)

- Hosts WebView2 inside VBA/Excel UserForms.
- Primary entry point is `CreateFromUserform`, which mounts the control on the form’s client area (same model `wuiUI` should follow).
- `CreateFromFrame` remains available when you need the WebView inside a specific `MSForms.Frame`.
- Renders HTML via `.Html` and supports script execution.
- Supports host object injection and request interception for advanced scenarios.

### `stdICallable` (available)

- Defines callback contract (`Run`, `RunEx`, `Bind`) for event integration.
- Enables passing callback objects into fluent UI builders.

## Proposed `wuiUI` API Direction

`wuiUI` should unify the above into one cohesive surface:

- Component-style, fluent element creation.
- Use `CreateChild("<tag>")` for standard HTML element behavior.
- First-class callback arguments on controls/elements.
- Automatic sender binding in event-capable controls (for example `CreateButton`).
- Mount/update cycle managed by VBA with scheduled re-renders.
- Optional node-level patch behavior to avoid full page rewrites for small changes.
- Host the WebView from a **UserForm** via `CreateFromUserform` (mirroring `stdWebView.CreateFromUserform`). Use a frame-based factory only when you deliberately embed inside `MSForms.Frame`.

Example direction:

```vb
Dim ui As wuiUI
Set ui = wuiUI.CreateFromUserform(host:=UserForm1)

With ui.CreateChild("div", Array("id", "root"))
  .CreateLiteral "Employee Dashboard"

  With .CreateButton( _
    text:="Refresh", _
    callback:=stdCallback.CreateFromModule("mUI", "OnRefreshClick") _
  )
    .ClassName = "btn btn-primary"
  End With
End With

ui.Render
```

Another fluent composition style:

```vb
With myDiv
  With .CreateButton(callback:=stdCallback.CreateFromModule("mUI", "OnSave"))
    .CreateLiteral "Save"
  End With
End With
```

## Runtime Model (Target)

1. Build UI tree in VBA (`wuiUI` + `wuiElement`, backed by `stdHTML`).
2. Render/mount into WebView (`stdWebView`), usually created with `CreateFromUserform` on the hosting form.
3. User actions trigger callbacks (`stdICallable` instances) with sender metadata.
4. VBA mutates state and calls `RequestRender` (batched/scheduled) instead of immediate recursive renders.
5. Render loop coalesces multiple updates and can choose full render or targeted node updates.

This keeps business logic, state, and event orchestration in VBA, with HTML/WebView as the rendering target.

## Event And State Model (Proposed)

- `wuiElement` nodes receive stable internal IDs on creation.
- Controls like `CreateButton` auto-bind sender context (for example `senderId`) into callbacks.
- Callback handlers should mutate state, then call `RequestRender`.
- Render loop should guard reentrancy (`isRendering`) and coalesce repeated render requests.
- For collections/lists, callbacks should use stable item IDs (not visual indexes).

## HTMX Positioning

HTMX-inspired ideas (attribute-driven interactions, partial updates) are useful references, but `wuiUI` is not intended to be an HTMX wrapper.  
The framework is optimized for VBA ergonomics first, with its own API and lifecycle.

## Non-Goals (Initial)

- Reproducing the full React runtime model in JavaScript.
- Building a JS-first framework where VBA becomes a thin transport layer.
- Tight coupling to HTMX conventions as the primary authoring model.

## Suggested First Milestones

1. `wuiUI.CreateFromUserform(host)` (and optionally `CreateFromFrame` for frame embedding), plus `wuiUI.Render`.
2. `CreateChild("<tag>")` as the default builder plus callback-aware controls (`CreateButton`, `CreateInput`, etc.).
3. Event dispatch bridge based on `stdICallable`.
4. Add `RequestRender` with coalescing and reentrancy guards.
5. Minimal update strategy (`Render` + targeted replace for known nodes).
6. Examples showing component-like composition in pure VBA.

## Custom Component Example: Accordion (Concept)

This example shows a reusable "custom component" shape in pure VBA.  
`RenderAccordion` behaves like a component function: it accepts a parent node, input data ("props"), and render context (`ui`), then renders itself and wires callbacks.

```vb
Private AccordionOpenById As Object 'Scripting.Dictionary<id, Boolean>

Public Sub ShowAccordionDemo()
  Set AccordionOpenById = CreateObject("Scripting.Dictionary")

  Dim ui As wuiUI
  Set ui = wuiUI.CreateFromUserform( _
    host:=UserForm1, _
    renderer:=stdCallback.CreateFromModule("mAccordionDemo", "RenderAccordionDemo") _
  )
  ui.RequestRender
End Sub

'Root render callback invoked by framework render loop
Public Sub RenderAccordionDemo(ByVal ui As wuiUI)
  ui.Clear

  Dim sections As Collection
  Set sections = New Collection
  sections.Add CreateAccordionSection(1, "What is wuiUI?", "A VBA-first UI framework concept on top of stdHTML + stdWebView.")
  sections.Add CreateAccordionSection(2, "Can I create custom components?", "Yes. Compose reusable render functions and bind callbacks.")
  sections.Add CreateAccordionSection(3, "How do updates work?", "Callbacks mutate VBA state, then call RequestRender.")

  With ui.CreateChild("div", Array("class", "accordion-demo"))
    .CreateChild("h2").CreateLiteral "Accordion Component"
    Call RenderAccordion(.Self, sections, ui)
  End With

  ui.Render
End Sub

'Custom component renderer (component-like API surface)
Public Sub RenderAccordion(ByVal parent As wuiElement, ByVal sections As Collection, ByVal ui As wuiUI)
  Dim section As Object
  For Each section In sections
    Dim sectionId As Long
    sectionId = CLng(section("id"))

    Dim isOpen As Boolean
    If AccordionOpenById.Exists(CStr(sectionId)) Then
      isOpen = CBool(AccordionOpenById(CStr(sectionId)))
    End If

    With parent.CreateChild("div", Array("class", "accordion-item", "data-id", sectionId))
      .CreateButton( _
        text:=CStr(section("title")), _
        callback:=stdCallback.CreateFromModule("mAccordionDemo", "OnAccordionToggle").Bind(sectionId) _
      )

      If isOpen Then
        .CreateChild("div", Array("class", "accordion-panel")).CreateLiteral CStr(section("content"))
      End If
    End With
  Next section
End Sub

Public Sub OnAccordionToggle(ByVal sectionId As Long, ByVal ctx As wuiEventContext)
  Dim key As String
  key = CStr(sectionId)

  Dim nextValue As Boolean
  If AccordionOpenById.Exists(key) Then
    nextValue = Not CBool(AccordionOpenById(key))
  Else
    nextValue = True
  End If

  AccordionOpenById(key) = nextValue
  ctx.UI.RequestRender
End Sub

Private Function CreateAccordionSection(ByVal id As Long, ByVal title As String, ByVal content As String) As Object
  Dim item As Object
  Set item = CreateObject("Scripting.Dictionary")
  item("id") = id
  item("title") = title
  item("content") = content
  Set CreateAccordionSection = item
End Function
```

Why this maps well to custom components:

- `RenderAccordion(...)` encapsulates structure + event wiring in one reusable unit.
- `sections` is equivalent to component props.
- `AccordionOpenById` is local module state for the component instance.
- `OnAccordionToggle` acts like an event handler that updates state, then calls `RequestRender`.

## Sample Todo App (Concept)

This example shows the intended shape of a tiny `wuiUI` todo app where VBA owns state and event handling.

```vb
Private Todos As Collection
Private TodoNextId As Long

Public Sub ShowTodoApp()
  Set Todos = New Collection
  TodoNextId = 0
  AddTodo "Write docs"
  AddTodo "Ship v1"

  Dim ui As wuiUI
  Set ui = wuiUI.CreateFromUserform( _
    host:=UserForm1, _
    renderer:=stdCallback.CreateFromModule("mTodoApp", "RenderTodoApp") _
  )
  ui.RequestRender
End Sub

'Invoked by framework render loop
Public Sub RenderTodoApp(ByVal ui As wuiUI) 
  ui.Clear

  With ui.CreateChild("div", Array("id", "todo-app"))
    .CreateChild("h2").CreateLiteral "Todo App"

    With .CreateChild("div", Array("class", "todo-input-row"))
      Dim textIn As wuiElement
      Set textIn = .CreateInput(placeholder:="Add task...")
      .CreateButton( _
        text:="Add", _
        callback:=stdCallback.CreateFromModule("mTodoApp", "OnAddTodo").Bind(textIn) _
      )
    End With

    Dim todo As Object
    For Each todo In Todos
      With .CreateChild("div", Array("class", "todo-row"))
        .CreateLiteral CStr(todo("text"))
        .CreateButton( _
          text:="Done", _
          callback:=stdCallback.CreateFromModule("mTodoApp", "OnCompleteTodo").Bind(CLng(todo("id"))) _
        )
      End With
    Next todo
  End With

  ui.Render 'Framework may optimize to node-level patches internally
End Sub

Public Sub OnAddTodo(ByVal textIn As wuiElement, ByVal ctx As wuiEventContext)
  Dim newText As String
  newText = Trim$(textIn.Text)
  If Len(newText) = 0 Then Exit Sub

  AddTodo newText
  textIn.Text = vbNullString
  ctx.UI.RequestRender
End Sub

Public Sub OnCompleteTodo(ByVal todoId As Long, ByVal ctx As wuiEventContext)
  RemoveTodoById todoId
  ctx.UI.RequestRender
End Sub

Private Sub AddTodo(ByVal text As String)
  Dim todo As Object
  Set todo = CreateObject("Scripting.Dictionary")
  TodoNextId = TodoNextId + 1
  todo("id") = TodoNextId
  todo("text") = text
  Todos.Add todo
End Sub

Private Sub RemoveTodoById(ByVal todoId As Long)
  Dim i As Long
  For i = Todos.Count To 1 Step -1
    If CLng(Todos(i)("id")) = todoId Then
      Todos.Remove i
      Exit Sub
    End If
  Next i
End Sub
```

The main idea is simple: callbacks receive sender/context data, mutate VBA state, and call `RequestRender` so updates are batched and safe.

In a generic `wuiUI` implementation, `RequestRender` should not hardcode a module procedure. Instead, the app registers a render callback (for example via `CreateFromUserform(..., renderer:=...)` or `SetRenderer`), and the framework render loop invokes that callback with the `ui` instance.

