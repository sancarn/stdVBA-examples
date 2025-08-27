# Document Generator

This example demonstrates how to generate documents from a template document and a tabular data source.

## Requirements

* stdVBA
    * stdArray
    * stdCallback
    * stdCOM
    * stdICallable
    * stdLambda
    * stdReg
    * stdShell
    * stdWindow
* Currently only works on Windows OS

## Usage


## High Level Process

```mermaid
flowchart TD
    A[Start] --> B[Load Lookups from Admin Table]
    B --> C[Select Target Factory<br/>Excel or PowerPoint Injector]
    C --> D[Create Injector Instance via TargetLambda]
    D --> E[Get Formula Bindings from Injector]
    E --> F[Loop over Bindings]
    F --> G[Compile Binding Lambda via genLambdaEx]
    G --> F
    F -->|Bindings ready| H[Prepare Source Table via stdTable]
    H --> I[Loop over Rows in Source]
    I --> J[Initialise Target Document via Injector]
    J --> K[Loop over Bindings for Row]
    K --> L[Evaluate Lambda on Row + Target]
    L --> M[Run Setter to Update Target]
    M --> K
    K --> N[Run AfterUpdate Lambda]
    N --> O[Cleanup Target Document via Injector]
    O --> I
    I -->|All docs generated| P[Finish]
```

## Project Structure

```mermaid
flowchart LR
    %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
    linkStyle default interpolate linear

    subgraph BaseLibraries[stdVBA]
        SL[stdLambda]
        SCB[stdCallback]
        SE[stdEnumerator]
        SR[stdRegex]
        SJ[stdJSON]

        subgraph stdExtensions[stdVBA Examples]
            ST[stdTable]
        end
    end

    subgraph InjectorFramework[Generic Injector Framework]
        GM[genMain]
        GI[genIInjector Interface]
        GX[genInjectorExcel]
        GP[genInjectorPowerPoint]
        GL[genLambdaEx]

        
    end
    
    %% Main flow
    GM --> GI
    GM --> GX
    GM --> GP
    GM --> GL
    GM --> ST

    %% Injector implementations
    GX --> GI
    GP --> GI

    %% Dependencies
    GX --> SL
    GX --> SCB
    GP --> SL
    GP --> SCB
    GL --> SL
    GL --> SCB
    GL --> SE
    GL --> SR

    %% stdTable dependencies
    ST --> SE
    ST --> SCB
    ST --> SJ
```