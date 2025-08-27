# Notepad Extract and Close

This example demonstrates how to extract all text from all open Notepad windows and close them afterwards.

## Requirements

* stdVBA
    * stdProcess
    * stdWindow
    * stdAcc
    * stdLambda
    * stdICallable

## High Level Process

```mermaid
flowchart TD
    A[Start executeMain or executeNoQuit] --> B[Clear Sheet1 and Write Headers]
    B --> C[Find Notepad Processes via stdProcess and stdLambda]

    C --> D[Loop over Processes]
    D --> E[Get Windows for Process via stdWindow]
    E --> F[Get Main Accessibility Object via stdAcc]
    F --> G[Activate All Tabs ROLE_PAGETAB]

    G --> H[Loop over Edit Windows Class = RichEditD2DPT]
    H --> I[Get Edit Accessibility Object and Extract Text]
    I --> J[Write Title and Text to Excel Sheet]
    J --> H

    H -->|No edit windows left, Next Process| D
    D -->|No more processes left| K[Optionally Quit Notepad Processes]
    K --> L[Finish]
```

## Project Structure

```mermaid
flowchart LR
    subgraph BaseLibraries[stdVBA]
        SP[stdProcess]
        SW[stdWindow]
        SA[stdAcc]
        SL[stdLambda]
    end

    subgraph MainModule[main Module]
        M[Notepad Extract and Close loop]
    end

    %% Dependencies
    M --> SP
    M --> SW
    M --> SA
    M --> SL
```