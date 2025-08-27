# Conditional Formatting using stdLambda

A very simple demo, one of the sheets contains a lookup table:

![definition table](./resources/image.png)

This list object is then iterated through whenever a cell changes on `ConditionalFormattingSheet`:

![demo](./resources/demo.gif)

## High Level Process

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    linkStyle default interpolate linear

    A[BulkApplyFormatting] --> B[TargetedApplyFormatting on shTest.UsedRange]
    B --> C[Load ConditionalFormatting Table<br/>via stdEnumerator]
    C --> D[Loop over Rules<br/>Compile Lambda Expressions]
    D --> E[Loop over Target Cells]
    E --> F[If Cell has Value → Run Rules]
    F --> G[Find First Matching Rule via Lambda]
    G -->|Match Found| H[Apply Interior Color from Rule]
    G -->|No Match| I[Clear Cell Interior Formatting]
    H --> E
    I --> E
    E --> J[Finish]
```

## Project Structure

```mermaid
flowchart LR
    subgraph BaseLibraries[stdVBA Utilities]
        SE[stdEnumerator]
        SL[stdLambda]
    end

    subgraph ExcelTables[Excel Tables]
        TCF[ConditionalFormatting ListObject]
    end

    subgraph Module[ConditionalFormattingEx Module]
        CF[ConditionalFormattingEx]
    end

    %% Dependencies
    CF --> SE
    CF --> SL
    CF --> TCF
```