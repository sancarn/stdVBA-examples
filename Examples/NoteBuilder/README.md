# NoteBuilder

A simple note-builder driven by stdVBA. 

# Spec

We define some questions to be asked to the user

| TemplateName   | Type     | Userform-Description    | dropdown-choices                       | checkbox-yes-text                       | checkbox-no-text |
| -------------- | -------- | ----------------------- | -------------------------------------- | --------------------------------------- | ---------------- |
| issue          | Dropdown | Issue Type              | flooding;smell;pollution               | N/A                                     | N/A              |
| cause          | Dropdown | Cause Type              | blockage;collapse;hydraulic incapacity | N/A                                     | N/A              |
| custVulnerable | Checkbox | Is customer vulnerable? | N/A                                    |  Customer is vulnerable.                |                  |
| job-raised     | Dropdown | Job raised              | jetting;pipe repair                    | N/A                                     | N/A              |
| capex-required | Checkbox | Is CAPEX required?      | N/A                                    |  CAPEX Required. Raised on risk system. |                  |

We present a userform to the user to select answers to the questions, and generate some note out the backend from a template already provided in a textbox.

Inspiration: https://www.reddit.com/r/vba/comments/1ixxv6u/is_there_something_we_can_just_pay_someone/

## High Level Process

```mermaid
flowchart TD
    A[Open Questionnaire Form] --> B[Load UserformElements Table into rows via stdEnumerator]

    %% Loop for building UI
    B --> C[Loop over Rows]
    C --> D[Create UI Controls via stdUIElement<br/>Label + Dropdown or Checkbox or Textbox]
    D --> C

    C --> F>User Selects Answers in Form and submits]
    F --> G[Load Note Template from shTemplate]

    %% Loop for replacing placeholders
    G --> H[Loop over Rows]
    H --> I[Replace Placeholder with User Input<br/>via protReduceRow]
    I --> H

    H --> J[Assemble Final Note Text]
    J --> K[Copy Note to Clipboard via stdClipboard]
    K --> L[End]
```

## Project Structure

```mermaid
flowchart LR
    linkStyle default interpolate linear
    
    subgraph BaseLibraries[stdVBA Utilities]
        SE[stdEnumerator]
        SCB[stdCallback]
        SUI[stdUIElement]
        SCL[stdClipboard]
    end

    subgraph ExcelTables1[Excel UI Build table]
        T1[UserformElements Table]
    end
    subgraph ExcelTables2[Excel Text Template]
        T2[NoteTemplate Shape]
    end

    subgraph NoteBuilder[Questionnaire Form]
        QF[Questionnaire]

        %% Split into two flows
        QF --> UI[UI Build]
        QF --> DH[Data Handling]
    end

    %% UI Build dependencies
    UI --> SE
    UI --> SCB
    UI --> SUI
    UI --> T1

    %% Data Handling dependencies
    DH --> SE
    DH --> SCB
    DH --> SCL
    DH --> T2
```