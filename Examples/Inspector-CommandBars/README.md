# CommandBar Inspector

CommandBar Inspector is a utility designed for VBA developers and Office power users to explore and inspect the internal `CommandBar` structure of Microsoft Office applications, such as Excel. It provides quick access to `CommandBar` IDs and labels for use in customization, automation, or troubleshooting.

## Features

* 🔍 Searchable Interface: Filter `CommandBars` and controls by name or ID.
* 📋 Copy MSO ID: Easily copy the internal MSO control ID for Ribbon or CommandBar customization.
* 🖱 Execute via Double-Click: Double-click a row to execute the corresponding control.
* 🖨 Print: Copy (ctrl+c) or print VBA to execute the command bar control to the clipboard.
* 📄 Support for Multiple Contexts: Switch between `Application.CommandBars` and `Application.VBE.CommandBars`!

## How to Use

1. Open the workbook
2. Ensure macros are enabled
3. Press the button on the main sheet.
4. Search/Find a command bar you want to use

## High Level Process

```mermaid
flowchart TD
    A[Open CommandBar Inspector Form] --> B[Initialize InkCollector for MouseWheel Support]
    B --> C[Populate ModeSwitcher with App and VBE]
    C --> D[User Selects Mode]
    D --> E[Loop over CommandBars and Controls]
    E --> F[Build Dictionary for Each Control<br/>ID, Parent, Caption, Command String]
    F --> G[Push Control Dictionaries into stdArray]
    G --> H[QueriedControls = All Controls]
    H --> I[Update ListBox with ID, Parent, Caption]

    I --> J[User Types in SearchBox]
    J --> K[Filter Controls via stdLambda Query<br/>on sanitizedName]
    K --> I

    I --> L[User Selects Control in ListBox]
    L --> M[Copy MSO ID to Clipboard<br/>via stdClipboard]
    L --> N[Print Execute Command to Immediate Window]
    L --> O[Double-Click → Execute Control]
    L --> P[Ctrl+C → Copy Execute Command to Clipboard]
```

## Project Structure

```mermaid
flowchart LR
    subgraph BaseLibraries[stdVBA Utilities]
        SA[stdArray]
        SL[stdLambda]
        SCB[stdClipboard]
        SW[stdWindow]
        SA2[stdAcc]
    end

    subgraph UIHelpers[UI Helpers]
        DICT[Dictionary - FastDictionary]
    end

    subgraph InspectorForm[InspectCommandbars Form]
        IC[InspectCommandbars]
        M1[Module1.ShowInspector]
    end

    %% Dependencies
    M1 --> IC
    IC --> SA
    IC --> SL
    IC --> SCB
    IC --> SW
    IC --> SA2
    IC --> DICT
```