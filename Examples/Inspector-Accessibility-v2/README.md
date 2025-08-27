<!--
    {
        "description": "Accessibility Inspector",
        "tags":["ui", "window", "automation", "embedding"],
        "deps":["stdAcc", "stdProcess", "stdWindow", "stdICallable"]
    }
-->

# Accessibility Inspector

While using `stdAcc` it is often useful to be able to obtain the accessibility information at the cursor. This can help you find elements to further investigate the accessibility tree. This example provides a utility application which can be used to:

* Pinpoint element attributes to assist during automation.
* Allows setting of `accValue`, typically useful to test setting fields with information.
* Allows execution of `DoDefaultAction`.


![inspector](./docs/InspectorTutorial.png)

## Requirements

* [stdVBA](http://github.com/sancarn/stdVBA)
    * stdAcc
    * stdCallback
    * stdClipboard
    * stdICallable
    * stdImage
    * stdLambda
    * stdProcess
    * stdShell
    * stdWindow
* tvTree
* uiVBA
    * uiElement
    * uiMessagable
* Currently only works on Windows OS

## Usage

Open xlsm and click "Show Accessibility Inspector"!

Navigate the treeview to insect the accessibility information of desktop windows.

## Roadmap

* [X] Extract basic accessibility information
* [X] Provide watchable cursor option.
* [X] Provide a temporary watchable cursor option ( 5 second timeout ).
* [X] Make form topmost
* [X] Code generation algorithm to generate stdAcc code for usage in user applications.
* [X] Search function to allow searching of accessibility tree.
* [X] Ability to only show visible elements.
* [X] Option to highlighting the selected accessibility element with a yellow rect.
* [ ] Option to find and display the hovered element within the accessibility tree.

## High Level Process

```mermaid
flowchart TD
    A[Open Accessibility Inspector Form] --> B[Initialize TreeControl and Property Panel]
    B --> C[Create tvTree with Root Element<br/>via tvAcc.CreateFromDesktop]
    C --> D[Add Fields to Property Panel<br/>via uiFields.AddField with stdLambda/stdCallback]

    %% Global actions
    C --> GA[User Can Search, Follow Mouse, or Refresh Tree]

    %% Tree interaction
    C --> E[User Expands Tree Node]
    E --> F[Retrieve Children via tvAcc.Children]
    F --> G[Loop over Children]
    G --> H[Add Child Nodes to Tree]
    H --> E

    %% Selection interaction
    D --> I[User Selects Element in Tree]
    I --> J[Update Property Panel via uiFields.UpdateSelection]
    J --> K[Display Properties: Identity, Name, Value, Role, States, etc.]
    K --> L[Optional Actions: DoDefaultAction, SetValue, Copy Selector, Highlight Rectangle]
```

## Project Structure

```mermaid
flowchart LR
    subgraph BaseLibraries[stdVBA Utilities]
        SL[stdLambda]
        SCB[stdCallback]
        SW[stdWindow]
        SP[stdProcess]
        SA[stdAcc]
        SCL[stdClipboard]
    end

    subgraph UIHelpers[UI Helpers]
        UF[uiFields]
        UE[uiElement]
        UM[uiIMessagable]
    end

    subgraph TreeHelpers[Tree Helpers]
        TT[tvTree]
        TA[tvAcc]
    end

    subgraph Inspector[AccessibilityInspector Form]
        AI[AccessibilityInspector]
    end

    %% Inspector dependencies
    AI --> TT
    AI --> UF
    AI --> SL
    AI --> SCB
    AI --> SW
    AI --> SP
    AI --> SA
    AI --> SCL

    %% Tree dependencies
    TT --> SL
    TT --> SCB
    TT --> TA

    %% tvAcc depends on stdAcc
    TA --> SA

    %% uiFields internals
    UF --> UE
    UF --> UM
    UF --> SL
    UF --> SCB
```