# Highlight Selected Cells

This example demonstrates how to highlight the selected cells in a spreadsheet using `stdWindow.CreateHighlightRect`.

## Requirements

* stdVBA
    * stdWindow
    * stdICallable
* Currently only works on Windows OS

## Usage

Open the attached xlsm demo, and highlight the cells in the spreadsheet.

## High Level Process

```mermaid
flowchart TD
    A[User Selects a Cell or Range] --> C[Worksheet_SelectionChange Event Fires]
    C --> D[Convert Range to Screen Coordinates<br/>via ObjRect using PointsToScreenPixelsX/Y]
    D --> E[Create Highlight Rectangle<br/>via stdWindow.CreateHighlightRect]
    E --> F[Overlay Window Highlights Selected Range]
```

## Project Structure

```mermaid
flowchart LR
    subgraph BaseLibraries[stdVBA]
        SW[stdWindow]
    end

    subgraph ExcelIntegration[Excel Sheet Code]
        S1[Sheet1]
    end

    %% Dependencies
    S1 --> SW
```