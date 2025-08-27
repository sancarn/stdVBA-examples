# Registry Viewer

## Usage

1. Click `Show Registry Viewer` button
2. Search the tree.

![_](./res/Process.png)

You can also copy the stdVBA code for selecting the registry key.

## High Level Process

```mermaid
flowchart TD
    A[Open Registry Viewer Form] --> B[Initialize Roots Collection HKCU, HKLM, HKCR, HKU]
    B --> C[Create tvTree with Roots using stdLambda for ID, Name, Children]
    C --> D[Populate TreeView with Root Keys]
    D --> E[User Expands Node]
    E --> F[Loop over Subkeys and Items via stdReg.Keys and stdReg.Items]
    F --> G[Add Child Nodes to TreeView]
    G --> E

    D --> H[User Selects Node]
    H --> I[Set SelectedEntry = stdReg Object]
    I --> J[Update Address Bar with Path]
    J --> K[Clear ListView and Populate with Items]
    K --> L[Loop over Registry Items Add Name, Type, Value]
    L --> K

    H --> M[Context Menu: Copy stdVBA Code]
    M --> N[Copy stdReg.CreateFromKey to Clipboard via stdClipboard]
```

## Project Structure

```mermaid
flowchart LR
    subgraph BaseLibraries[stdVBA Utilities]
        SL[stdLambda]
        SCB[stdClipboard]
    end

    subgraph RegistryCore[Registry Core]
        SR[stdReg]
    end

    subgraph UIHelpers[UI Helpers]
        TV[tvTree]
    end

    subgraph RegistryViewerForm[RegistryViewer Form]
        RV[RegistryViewer]
    end

    %% Dependencies
    RV --> TV
    RV --> SR
    RV --> SL
    RV --> SCB

    TV --> SL
    TV --> SR
```